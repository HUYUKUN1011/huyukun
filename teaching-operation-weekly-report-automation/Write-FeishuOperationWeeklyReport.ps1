param(
  [ValidateSet("check", "dry-run", "write")]
  [string]$Mode = "dry-run",

  [string]$ConfigPath = ".\operation_weekly_report_automation.config.json",
  [string]$DraftPath = "",
  [string]$DraftJson = "",
  [string]$TargetDate = ""
)

$ErrorActionPreference = "Stop"

function Write-JsonResult {
  param([object]$Value)
  $Value | ConvertTo-Json -Depth 20
}

function Get-RequiredEnv {
  param([string]$Name)
  $value = [Environment]::GetEnvironmentVariable($Name, "Process")
  if (-not $value) {
    $value = [Environment]::GetEnvironmentVariable($Name, "User")
  }
  if (-not $value) {
    $value = [Environment]::GetEnvironmentVariable($Name, "Machine")
  }
  if (-not $value) {
    throw "Missing required environment variable: $Name"
  }
  return $value
}

function Invoke-FeishuApi {
  param(
    [ValidateSet("GET", "POST", "PATCH", "DELETE")]
    [string]$Method,
    [string]$Path,
    [object]$Body = $null,
    [string]$TenantToken = ""
  )

  $headers = @{}
  if ($TenantToken) {
    $headers["Authorization"] = "Bearer $TenantToken"
  }

  $uri = "https://open.feishu.cn/open-apis$Path"
  try {
    if ($null -ne $Body) {
      $json = $Body | ConvertTo-Json -Depth 30 -Compress
      return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -ContentType "application/json; charset=utf-8" -Body $json
    }
    return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
  } catch {
    $detail = $_.Exception.Message
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
      $detail = $_.ErrorDetails.Message
    }
    throw "Feishu API $Method $Path failed: $detail"
  }
}

function Get-TenantAccessToken {
  param([string]$AppId, [string]$AppSecret)
  $response = Invoke-FeishuApi -Method POST -Path "/auth/v3/tenant_access_token/internal" -Body @{
    app_id = $AppId
    app_secret = $AppSecret
  }
  if ($response.code -ne 0 -or -not $response.tenant_access_token) {
    throw "Failed to get tenant access token: $($response | ConvertTo-Json -Depth 10 -Compress)"
  }
  return $response.tenant_access_token
}

function Resolve-WikiNode {
  param([string]$WikiToken, [string]$TenantToken)
  $encodedToken = [Uri]::EscapeDataString($WikiToken)
  $response = Invoke-FeishuApi -Method GET -Path "/wiki/v2/spaces/get_node?token=$encodedToken" -TenantToken $TenantToken
  if ($response.code -ne 0) {
    throw "Failed to resolve wiki node: $($response | ConvertTo-Json -Depth 10 -Compress)"
  }
  if ($response.data.node) {
    return $response.data.node
  }
  if ($response.data) {
    return $response.data
  }
  throw "Wiki node response did not include node data."
}

function Get-DocxDocument {
  param([string]$DocumentId, [string]$TenantToken)
  $encodedId = [Uri]::EscapeDataString($DocumentId)
  $response = Invoke-FeishuApi -Method GET -Path "/docx/v1/documents/$encodedId" -TenantToken $TenantToken
  if ($response.code -ne 0) {
    throw "Failed to read docx document: $($response | ConvertTo-Json -Depth 10 -Compress)"
  }
  if ($response.data.document) {
    return $response.data.document
  }
  return $response.data
}

function Get-DocxBlocks {
  param([string]$DocumentId, [string]$TenantToken)
  $encodedId = [Uri]::EscapeDataString($DocumentId)
  $blocks = @()
  $pageToken = ""
  do {
    $path = "/docx/v1/documents/$encodedId/blocks?page_size=500"
    if ($pageToken) {
      $path += "&page_token=$([Uri]::EscapeDataString($pageToken))"
    }
    $response = Invoke-FeishuApi -Method GET -Path $path -TenantToken $TenantToken
    if ($response.code -ne 0) {
      throw "Failed to list docx blocks: $($response | ConvertTo-Json -Depth 10 -Compress)"
    }
    $items = @()
    if ($response.data.items) {
      $items = @($response.data.items)
    } elseif ($response.data.blocks) {
      $items = @($response.data.blocks)
    }
    $blocks += $items
    $pageToken = $response.data.page_token
  } while ($pageToken)
  return @($blocks)
}

function Get-BlockText {
  param([object]$Block)

  $textContainers = @(
    $Block.text,
    $Block.heading1,
    $Block.heading2,
    $Block.heading3,
    $Block.heading4,
    $Block.heading5,
    $Block.heading6,
    $Block.bullet,
    $Block.ordered
  ) | Where-Object { $null -ne $_ }

  $parts = @()
  foreach ($container in $textContainers) {
    if ($container.elements) {
      foreach ($element in @($container.elements)) {
        if ($element.text_run -and $null -ne $element.text_run.content) {
          $parts += [string]$element.text_run.content
        } elseif ($element.mention_user -and $element.mention_user.name) {
          $parts += [string]$element.mention_user.name
        }
      }
    }
  }
  return (($parts -join "") -replace "[\u200b\u200c\u200d\ufeff]", "").Trim()
}

function New-TextBlock {
  param([string]$Content)
  if ([string]::IsNullOrWhiteSpace($Content)) {
    $Content = " "
  }
  return @{
    block_type = 2
    text = @{
      elements = @(
        @{
          text_run = @{
            content = $Content
            text_element_style = @{}
          }
        }
      )
      style = @{}
    }
  }
}

function New-BulletBlock {
  param([string]$Content)
  return @{
    block_type = 12
    bullet = @{
      elements = @(
        @{
          text_run = @{
            content = $Content
            text_element_style = @{}
          }
        }
      )
      style = @{}
    }
  }
}

function Get-NextMondayTitle {
  param([string]$DateOverride)
  if ($DateOverride) {
    $date = [DateTime]::Parse($DateOverride)
  } else {
    $now = Get-Date
    $daysUntilMonday = ([int][DayOfWeek]::Monday - [int]$now.DayOfWeek + 7) % 7
    if ($daysUntilMonday -eq 0) {
      $daysUntilMonday = 7
    }
    $date = $now.Date.AddDays($daysUntilMonday)
  }
  return ("{0}/{1}/{2}" -f $date.Year, $date.Month, $date.Day)
}

function Import-Draft {
  param([string]$DraftPath, [string]$DraftJson)
  if ($DraftJson) {
    return ($DraftJson | ConvertFrom-Json)
  }
  if ($DraftPath) {
    return (Get-Content -Raw -Encoding UTF8 $DraftPath | ConvertFrom-Json)
  }
  return $null
}

function Get-SectionItems {
  param([object]$Draft, [object]$Section)
  if (-not $Draft) {
    return @()
  }

  $keys = @($Section.title, $Section.owner, $Section.alias) | Where-Object { $_ }

  foreach ($key in $keys) {
    if ($Draft.PSObject.Properties.Name -contains $key) {
      $value = $Draft.$key
      if ($value -is [array]) {
        return @($value | ForEach-Object { [string]$_ })
      }
      if ($value.items) {
        return @($value.items | ForEach-Object { [string]$_ })
      }
      if ($value.confirmed_items) {
        return @($value.confirmed_items | ForEach-Object { [string]$_ })
      }
      if ($value.confirmed) {
        return @($value.confirmed | ForEach-Object { [string]$_ })
      }
      if ($value.modified) {
        return @($value.modified | ForEach-Object { [string]$_ })
      }
      if ($value) {
        return @([string]$value)
      }
    }
  }

  if ($Draft.sections) {
    foreach ($entry in @($Draft.sections)) {
      $entryKeys = @($entry.title, $entry.owner, $entry.alias) | Where-Object { $_ }
      if (@($entryKeys | Where-Object { $keys -contains $_ }).Count -gt 0) {
        $items = @()
        foreach ($field in @("items", "confirmed_items", "confirmed", "modified")) {
          if ($entry.$field) {
            $items += @($entry.$field | ForEach-Object { [string]$_ })
          }
        }
        if ($items.Count -gt 0) {
          return @($items)
        }
        if ($entry.text) {
          return @([string]$entry.text)
        }
      }
    }
  }

  return @()
}

function Get-RootParentId {
  param([object[]]$Blocks, [object]$Document)
  $topBlock = @($Blocks | Where-Object { $_.parent_id } | Select-Object -First 1)[0]
  if ($topBlock -and $topBlock.parent_id) {
    return [string]$topBlock.parent_id
  }
  if ($Document.block_id) {
    return [string]$Document.block_id
  }
  if ($Document.document_id) {
    return [string]$Document.document_id
  }
  throw "Unable to infer root parent block id."
}

function Invoke-CreateChildren {
  param(
    [string]$DocumentId,
    [string]$ParentBlockId,
    [int]$Index,
    [object[]]$Children,
    [string]$TenantToken
  )

  if ($Children.Count -eq 0) {
    return $null
  }

  $encodedDoc = [Uri]::EscapeDataString($DocumentId)
  $encodedParent = [Uri]::EscapeDataString($ParentBlockId)
  $body = @{
    index = $Index
    children = @($Children)
  }
  $response = Invoke-FeishuApi -Method POST -Path "/docx/v1/documents/$encodedDoc/blocks/$encodedParent/children?document_revision_id=-1" -Body $body -TenantToken $TenantToken
  if ($response.code -ne 0) {
    throw "Failed to create docx blocks: $($response | ConvertTo-Json -Depth 20 -Compress)"
  }
  return $response
}

function Get-DateBlockIndexes {
  param([object[]]$BlockInfos)
  return @(
    for ($i = 0; $i -lt $BlockInfos.Count; $i++) {
      if ($BlockInfos[$i].text -match "^\d{4}/\d{1,2}/\d{1,2}$") {
        $i
      }
    }
  )
}

function Test-RangeHasContent {
  param([object[]]$BlockInfos, [int]$StartExclusive, [int]$EndExclusive)
  for ($i = $StartExclusive + 1; $i -lt $EndExclusive; $i++) {
    $text = $BlockInfos[$i].text
    if ($text -and -not ($text -match "^(\d{4}/\d{1,2}/\d{1,2}|上周会议重点待办)$")) {
      return $true
    }
  }
  return $false
}

$configFullPath = Resolve-Path $ConfigPath
$config = Get-Content -Raw -Encoding UTF8 $configFullPath | ConvertFrom-Json
$appId = Get-RequiredEnv "FEISHU_APP_ID"
$appSecret = Get-RequiredEnv "FEISHU_APP_SECRET"
$tenantToken = Get-TenantAccessToken -AppId $appId -AppSecret $appSecret

$wikiNode = Resolve-WikiNode -WikiToken $config.feishu_wiki_token -TenantToken $tenantToken
$documentId = [string]$wikiNode.obj_token
if (-not $documentId) {
  $documentId = [string]$config.feishu_obj_token
}
if ($config.feishu_obj_token -and $documentId -ne $config.feishu_obj_token) {
  throw "Wiki node obj_token '$documentId' does not match config feishu_obj_token '$($config.feishu_obj_token)'."
}

$document = Get-DocxDocument -DocumentId $documentId -TenantToken $tenantToken
$blocks = Get-DocxBlocks -DocumentId $documentId -TenantToken $tenantToken
$blockInfos = @(
  for ($i = 0; $i -lt $blocks.Count; $i++) {
    [PSCustomObject]@{
      index = $i
      block_id = [string]$blocks[$i].block_id
      parent_id = [string]$blocks[$i].parent_id
      block_type = $blocks[$i].block_type
      text = Get-BlockText $blocks[$i]
    }
  }
)

$targetTitle = Get-NextMondayTitle -DateOverride $TargetDate
$draft = Import-Draft -DraftPath $DraftPath -DraftJson $DraftJson
$sections = @($config.feishu_api_writeback.target_sections)
if ($sections.Count -eq 0) {
  throw "Config field feishu_api_writeback.target_sections must include the target weekly report sections."
}

$templateLines = @($targetTitle, "", "上周会议重点待办", "")
$sectionPlans = @()
foreach ($section in $sections) {
  $items = @(Get-SectionItems -Draft $draft -Section $section | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  $templateLines += [string]$section.title
  foreach ($item in $items) {
    $templateLines += "- $item"
  }
  $templateLines += ""
  $sectionPlans += [PSCustomObject]@{
    title = [string]$section.title
    owner = [string]$section.owner
    item_count = $items.Count
    items = $items
  }
}

$dateIndexes = @(Get-DateBlockIndexes -BlockInfos $blockInfos)
$targetIndex = -1
for ($i = 0; $i -lt $blockInfos.Count; $i++) {
  if ($blockInfos[$i].text -eq $targetTitle) {
    $targetIndex = $i
    break
  }
}

$latestDateIndex = if ($dateIndexes.Count -gt 0) { $dateIndexes[0] } else { -1 }
$parentId = Get-RootParentId -Blocks $blocks -Document $document
$operations = @()

if ($targetIndex -lt 0) {
  $insertIndex = if ($latestDateIndex -ge 0) { $latestDateIndex } else { [Math]::Min(1, $blocks.Count) }
  $children = @()
  foreach ($line in $templateLines) {
    if ($line -like "- *") {
      $children += New-BulletBlock -Content $line.Substring(2)
    } else {
      $children += New-TextBlock -Content $line
    }
  }
  $operations += [PSCustomObject]@{
    action = "create_next_week_template"
    parent_block_id = $parentId
    index = $insertIndex
    block_count = $children.Count
    lines = $templateLines
    children = $children
  }
} else {
  $nextDateIndex = $blockInfos.Count
  foreach ($idx in $dateIndexes) {
    if ($idx -gt $targetIndex) {
      $nextDateIndex = $idx
      break
    }
  }

  foreach ($sectionPlan in $sectionPlans) {
    if ($sectionPlan.item_count -eq 0) {
      continue
    }

    $sectionIndex = -1
    for ($i = $targetIndex + 1; $i -lt $nextDateIndex; $i++) {
      if ($blockInfos[$i].text -eq $sectionPlan.title) {
        $sectionIndex = $i
        break
      }
    }
    if ($sectionIndex -lt 0) {
      continue
    }

    $nextSectionIndex = $nextDateIndex
    foreach ($other in $sectionPlans) {
      for ($i = $sectionIndex + 1; $i -lt $nextDateIndex; $i++) {
        if ($blockInfos[$i].text -eq $other.title) {
          $nextSectionIndex = [Math]::Min($nextSectionIndex, $i)
          break
        }
      }
    }

    if (Test-RangeHasContent -BlockInfos $blockInfos -StartExclusive $sectionIndex -EndExclusive $nextSectionIndex) {
      $operations += [PSCustomObject]@{
        action = "skip_non_empty_section"
        section = $sectionPlan.title
        item_count = $sectionPlan.item_count
      }
      continue
    }

    $children = @($sectionPlan.items | ForEach-Object { New-BulletBlock -Content $_ })
    $operations += [PSCustomObject]@{
      action = "fill_empty_section"
      section = $sectionPlan.title
      parent_block_id = $parentId
      index = $sectionIndex + 1
      block_count = $children.Count
      lines = @($sectionPlan.items | ForEach-Object { "- $_" })
      children = $children
    }
  }
}

if ($Mode -eq "check" -or $Mode -eq "dry-run") {
  Write-JsonResult ([PSCustomObject]@{
    ok = $true
    mode = $Mode
    document_id = $documentId
    document_title = $document.title
    target_date_title = $targetTitle
    target_exists = ($targetIndex -ge 0)
    block_count = $blocks.Count
    operations = @($operations | ForEach-Object {
      [PSCustomObject]@{
        action = $_.action
        section = $_.section
        parent_block_id = $_.parent_block_id
        index = $_.index
        block_count = $_.block_count
        lines = $_.lines
      }
    })
    sections = $sectionPlans
  })
  exit 0
}

$writeResults = @()
foreach ($operation in $operations) {
  if ($operation.action -eq "skip_non_empty_section") {
    $writeResults += $operation
    continue
  }
  $response = Invoke-CreateChildren -DocumentId $documentId -ParentBlockId $operation.parent_block_id -Index ([int]$operation.index) -Children @($operation.children) -TenantToken $tenantToken
  $writeResults += [PSCustomObject]@{
    action = $operation.action
    section = $operation.section
    index = $operation.index
    block_count = $operation.block_count
    response = $response
  }
}

Write-JsonResult ([PSCustomObject]@{
  ok = $true
  mode = $Mode
  document_id = $documentId
  document_title = $document.title
  target_date_title = $targetTitle
  target_exists_before_write = ($targetIndex -ge 0)
  write_results = $writeResults
})
