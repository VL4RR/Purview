Connect-IPPSSession
Connect-ExchangeOnline
#Connect-MgGraph -Scopes "OrgSettings-AppsAndServices.Read.All", "Team.ReadBasic.All"

##Define HTML
$HTML = "C:\temp\CurrentComplianceReport.html"
$date = Get-Date -Format "yyyy-MM-dd"
Set-Content -Path $HTML -Value @"
<!DOCTYPE html>
<html>
<head>
<title>Microsoft Purview Current Compliance Posture</title>
<style>
body {font-family: "Segoe UI", sans-serif;background-color: #f8f9fa;color: #212529;margin: 20px;font-size: 17px;}
h1 {text-align: center;color: #000000;margin-bottom: 10px;}
h2 {background: #e3f2fd;color: #0d6efd;padding: 12px;border: 1px solid #ddd;border-left: 4px solid #0d6efd;border-radius: 4px;margin-top: 20px;cursor: pointer;position:relative;}
h2::after {content: " ▼";font-size: 0.9em;}
h2.active::after {content: " ▲";}
.section-content {display: none;margin-left: 20px;padding: 10px;}
h3 {background: #f8f9fa;color: #212529;padding: 12px;border: 1px solid #ddd;border-left: 3px solid #6c757d;border-radius: 4px;margin-top: 15px;cursor: pointer;position: relative;}
h3::after {content: " ▼";font-size: 0.9em;}
h3.active::after {content: " ▲";}
h3:hover {background: #e9ecef;}
.h3-content {display: none;margin-left: 25px;padding: 15px;background-color: #fff;border: 1px solid #ddd;border-radius: 4px;}
.tabs {display: flex;border-bottom: 2px solid #6c757d;margin-bottom: 20px;}
.tab-button {background: #f8f9fa;color: #212529;border:none;padding: 12px 20px;cursor:pointer;font-size:16px;font-weight:600;}
.tab-button:hover {background: #e9ecef;}
.tab-button.active {background: #0d6efd;color: white;}
.tab-content {display: none;}
.tab-content.active {display: block;}
.label-group {margin: 15px 0;border: 1px solid #dee2e6;border-radius: 4px;background-color: #f8f9fa;}
.label-group-header {color: #212529;background-color: #e9ecef;font-weight: 700;padding: 10px;border-left: 5px solid #6c757d;cursor:pointer;font-size:19px;}
.label-group-header:hover {background-color: #dee2e6;}
.label-group-content {padding: 15px;background-color: #fff;border: 1px solid #ddd;border-radius: 4px;}
h4 {color: #212529;font-weight: 600;margin: 10px 0 5px 0;}
table {width: 100%;border-collapse: collapse;margin: 15px 0;font-size: 16px;table-layout: fixed;}
th, td {border: 1px solid #ddd;padding: 9px 13px;text-align: left;font-size: 14px;word-wrap: break-word;vertical-align: top;}
th {background-color: #e9ecef;color: #212529;font-weight: 600;}
tr:nth-child(even) {background-color: #f8f9fa;}
tr:hover {background-color: #e9ecef;}
div {margin: 3px 0;font-size: 16px;color: #495057;}
hr {border: 0;border-top: 1px solid #dee2e6;margin: 20px 0;}
.controls {display: flex;gap: 10px;margin-bottom: 20px;padding: 10px;background: #f8f9fa;border-radius: 4px;}
.btn {background: #212529;color: white;border: none;padding: 8px 16px;border-radius: 4px;cursor: pointer;font-size: 14px;}
.btn:hover {background: #495057;}
.btn:secondary {background: #e9ecef;color: #212529;}
.btn:secondary:hover {background: #dee2e6;}
.condition-group {border: 1px solid #ccc; border-radius: 4px; padding: 10px; margin: 10px 0; background-color: #f9f9f9;}
.condition-operator-between {text-align: left; margin: 15px 0; font-weight: bold; color: #212529; font-size: 18px;}
.sit-group-operator {text-align: left; margin: 10px 0; font-weight: bold; color: #212529; font-size: 16px;}
.condition-operator-between-left {text-align: left; margin: 5px 0; font-weight: bold; color: #212529; font-size: 16px;}
.condition-not-operator {text-align: left; color: #212529; font-weight: bold; font-size: 16px; margin-left: 5px;}
.condition-group-heading {font-weight: bold; margin-bottom: 5px; display: flex; flex-wrap: wrap; align-items: center;}
.condition-item {margin: 5px 0; padding-left: 20px; /* indent for items */}
.group-depth-0 {border-color: #6c757d; background-color: #f8f9fa;}
.group-depth-1 {border-color: #6c757d; background-color: #f8f9fa;}
.group-depth-2 {border-color: #212529; background-color: #fff;}
.group-depth-3 {border-color: #0d6efd; background-color: #e3f2fd;}
.group-depth-4 {border-color: #6c757d; background-color: #f8f9fa;}
.advanced-rules-container {border: 1px dashed #6c757d; border-radius: 4px; padding: 10px; margin: 10px 0; background-color: #f9f9f9;}
</style>
</head>
<body>

<h1>Microsoft Purview Current Compliance Posture $date</h1>
<hr>

<div class="controls">
    <button class="btn" onclick="collapseAll()">Collapse All</button>
    <button class="btn btn-secondary" onclick="expandAll()">Expand All</button>
</div>

<div class="tabs">
    <button class="tab-button active" onclick="openTab(event,'InformationProtection')">Information Protection</button>
    <button class="tab-button" onclick="openTab(event,'Retention')">Retention</button>
    <button class="tab-button" onclick="openTab(event,'DLP')">DLP</button>
    <button class="tab-button" onclick="openTab(event,'TenantControls')">Tenant Controls</button>
</div>

"@ | Add-Content -Path $HTML

$polMode = @{
    "Enable"                   = "On"
    "TestWithoutNotifications" = "Simulation mode"
    "TestWithNotifications"    = "Simulation mode with policy tips"
    "Disable"                  = "Off"
}

$RecipientTypeConv = @{
    "2048"          = "Dynamic Distribution List"
    "524288"        = "Security Group"
    "256"           = "Distribution List"
    "1099511627776" = "M365 Group"
    "1024"          = "Mail-Enabled Security Group"
    "128"           = "User"
    "1"             = "User"
}

$allConditionsCondition = @{
    "ContentContainsSensitiveInformation" = "Content Contains:"
    "AccessScope"                         = "Content is shared:"
    "SenderIPRanges"                      = "Sender IP address is:"
    "RecipientDomainIs"                   = "Recipient domain is:"
    "ContentExtensionMatchesWords"        = "Attachment extension is:"
    "DocumentIsPasswordProtected"         = "Attachment is password protected"
    "DocumentIsUnsupported"               = "Any email attachment's content could not be scanned"
    "ProcessingLimitExceeded"             = "Any email attachment's content didn't complete scanning"
    "HeaderMatchesPatterns"               = "Header matches patterns:"
    "SubjectMatchesPatterns"              = "Subject matches patterns:"
    "AnyOfRecipientAddressContainsWords"  = "Recipient address contains words:"
    "FromAddressMatchesPatterns"          = "Sender address matches pattern"
    "SenderDomainIs"                      = "Sender domain is:"
    "SentToMemberOf"                      = "Recipient is a member of:"
    "From"                                = "Sender is:"
    "DocumentNameMatchesWords"            = "Attachment name matches words or phrases:"
    "ContentPropertyContainsWords"        = "Attachment property is:"
    "DocumentSizeOver"                    = "Attachment size equals or is greater than:"
    "SentTo"                              = "Recipient is:"

    "DocumentCreatedBy"                   = "Document created by:"
    "DocumentCreatedOnOrAfter"            = "Document creation date is on or after"
    "DocumentCreatedOnOrBefore"           = "Document creation date is on or before"
    "DocumentLastModifiedOnOrAfter"       = "Document modification date is on or after"
    "DocumentLastModifiedOnOrBefore"      = "Document modification date is on or before"
}

##########################################################################################################################################################


function labels {

    Write-Host "Gathering Sensitivity Labels"

    $actionLabels = @{
        "protectgroup"                                 = "Protect group"
        "privacy"                                      = "Privacy"
        "allowemailfromguestusers"                     = "Allow email from guest users"
        "allowaccesstoguestusers"                      = "Allow email from guest users"
        "disabled"                                     = "Disabled"
        "protectsite"                                  = "Protect site"
        "allowfullaccess"                              = "Allow full access"
        "allowlimitedaccess"                           = "Allow limited access"
        "blockaccess"                                  = "Block access"
        "externalsharingcontroltype"                   = "External sharing control type"
        "allowprivateteamstobediscoverableusingsearch" = "Allow private teams to be discoverable using search"
        "ischannelsharedwithexternaltenants"           = "Is channel shared with external tenants"
        "ischannelsharedwithsamelabelonly"             = "Is channel shared with same label only"
        "ischannelsharedwithprivateteamsonly"          = "Is channel shared with private teams only"
        "protectteamschannel"                          = "Protect teams channel"
        "applycontentmarking"                          = "Apply content marking"
        "text"                                         = "Text"
        "fontsize"                                     = "Font size"
        "fontcolor"                                    = "Font colour"
        "alignment"                                    = "Alignment"
        "margin"                                       = "Margin"
        "placement"                                    = "Placement"
        "applywatermarking"                            = "Apply watermarking"
        "fontname"                                     = "Font name"
        "applydynamicwatermarking"                     = "Apply dynamic watermarking"
        "encrypt"                                      = "Encrypt"
        "protectiontype"                               = "Protect type"
        "templateid"                                   = "Template ID"
        "templatearchived"                             = "Template archived"
        "linkedtemplateid"                             = "Linked template ID"
        "contentexpiredondateindaysornever"            = "User access to content expires"
        "offlineaccessdays"                            = "Allow offline access"
        "rightsdefinitions"                            = "Assign permissions to specific users and groups"
        "disableautomaticownerrights"                  = "Disable automatic owner rights"
        "dynamicwatermarkdisplay"                      = "Dynamic watermark"
        "protectteams"                                 = "Protect Teams"
        "islobbyrestrictionenforced"                   = "Control who can bypass the lobby"
        "ispresentersrestrictionenforced"              = "Control who can present"
        "iscopyrestrictionenforced"                    = "Prevent copying and forwarding of meeting content"
        "allowmeetingchat"                             = "Control meeting chat"
        "whocanrecord"                                 = "Control who can record and transcribe"
        "watermarkprotection"                          = "Control watermarks"
        "detectsensitivecontentduringscreensharing"    = "Detect Sensitive content during screen sharing"
        "lobbybypassscope"                             = "Who can bypass the lobby"
        "lobbybypassfordialinusersenabled"             = "People dialing in can bypass the lobby"
        "disablelobby"                                 = "Reject anyone who can't bypass the lobby"
        "allowedpresenters"                            = "Control who can present"
        "recordautomatically"                          = "Control whether meetings are recorded automatically"
        "isendtoendencryptionenabled"                  = "Control end-to-end encryption for meeting video and audio"
        "donotforward"                                 = "Do not forward"
        "promptuser"                                   = "Prompt user"
        "encryptonly"                                  = "Encrypt only"
        "protectionlevel"                              = "Protection level"
    }
    $valueLabels = @{
        "true"                            = "True"
        "false"                           = "False"
        "private"                         = "Private"
        "template"                        = "Template"
        "userdefined"                     = "User defined"
        "Organizer,Presenters"            = "Organizers and Presenters"
        "Organizer,Coorganizers"          = "Organizers and Co-organizers"
        "OrganizationAndFederated"        = "People in my org, trusted orgs, and guests"
        "Organization"                    = "People in my org and guests"
        "OrganizationExcludingGuests"     = "People in my org"
        "RoleIsPresenter"                 = "Let meeting organizers select specific people"
        "Organizer"                       = "Only organizers and co-organizers"
        "Invited"                         = "People who were invited"
        "public"                          = "Public"
        "ExistingExternalUserSharingOnly" = "Existing guests"
        "unspecified"                     = "Unspecified"
        "ExternalUserSharingOnly"         = "New and existing guests"
        "ExternalUserAndGuestSharing"     = "Anyone"
        "-1"                              = "Always"
    }

    '<h2>Sensitivity Labels</h2><div class="section-content">' | Add-Content -Path $HTML

    ##########################################################################################################################################################

    $Labels = Get-Label
    $script:Labels = $Labels

    $TopLevelLabels = @(
        $Labels |
        Where-Object {
            $_.IsLabelGroup -eq $true -or
            [string]::IsNullOrEmpty([string]$_.ParentId)
        } |
        Sort-Object Priority
    )

    Write-Host "Found $($Labels.Count) sensitivity labels ($($TopLevelLabels.Count) top-level labels)"

    $WriteLabel = {

        param (
            $Label
        )
        "<h3>Label - $($Label.DisplayName)&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;Priority: $($Label.Priority)</h3>" | Add-Content -Path $HTML
        "<div class='h3-content'><div style='background-color: #f8f9fa; padding: 10px; border-radius: 4px; margin: 0 0 8px 0; color: #212529;'><strong>Created by:</strong> $($Label.CreatedBy)<br><Strong>Last modified by: </Strong>$($Label.LastModifiedBy)<br><strong>Last modified:</strong> $(if ($Label.WhenChanged) { $Label.WhenChanged.ToString('yyyy-MM-dd') } else { '' })</div><div style='background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin: 0; color: #212529;'><strong>User Description:</strong> $($Label.Tooltip)</div>" | Add-Content -Path $HTML

        if ([string]::IsNullOrEmpty([string]$Label.Comment)) { '<div style="background-color: #e8f5e8; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;"><strong>Admin Description:</strong> No description provided</div>' | Add-Content -Path $HTML }
        else { "<div style='background-color: #e8f5e8; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Admin Description:</strong> $($Label.Comment)</div>" | Add-Content -Path $HTML }

        "<div style='background-color: #fff3cd; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Scope:</strong> $($Label.ContentType)</div>" | Add-Content -Path $HTML

        if ($Label.Capabilities) { '<div style="background-color: #f0e6ff; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;"><strong>Capabilities:</strong> {0}</div>' -f ($Label.Capabilities -join ', ') | Add-Content -Path $HTML }
        if ($null -eq $Label.LabelActions -or $Label.LabelActions.Count -eq 0) { '<p style="background-color: #e9ecef; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;">There are no Label actions</p>' | Add-Content -Path $HTML }
        else {
            foreach ($LabelAction in $Label.LabelActions) {

                try { $Action = $LabelAction | ConvertFrom-Json -ErrorAction Stop }
                catch { continue }

                $translatedType = $actionLabels[$Action.Type]
                if ($translatedType) { $Action.Type = $translatedType }

                if ($Action.SubType) {
                    $translatedSubType = $actionLabels[$Action.SubType]
                    if ($translatedSubType) { $Action.SubType = $translatedSubType }
                }

                "<h4 style='background-color: #d1ecf1; padding: 8px; border-radius: 4px; margin: 12px 0 8px 0;'><strong>$($Action.Type) $($Action.SubType)</strong></h4>" | Add-Content -Path $HTML
                '<table style="width: 100%; border-collapse: collapse; margin: 8px 0;"><tr><th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f8f9fa;">Configuration</th><th style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f8f9fa;">Value</th></tr>' | Add-Content -Path $HTML

                foreach ($Setting in $Action.Settings) {
                    $displayKey = $actionLabels[$Setting.Key]
                    if ($displayKey) {
                        $Setting.Key = $displayKey
                    }
                    # Translate the value if possible
                    $displayValue = $valueLabels[$Setting.Value]
                    if ($displayValue) {
                        $Setting.Value = $displayValue
                    }
                    # Special handling for rightsdefinitions (now displayed as "Rights")
                    if ($Setting.Key -eq "Assign permissions to specific users and groups" -and $Setting.Value) {
                        try {
                            $jsonValue = $Setting.Value | ConvertFrom-Json
                            if ($jsonValue) {
                                # Collect formatted strings for each item
                                $formattedItems = @()
                                if ($jsonValue -is [System.Array]) {
                                    foreach ($item in $jsonValue) {
                                        $identity = if ($item.Identity) { $item.Identity } else { "" }
                                        $rights = if ($item.Rights) { $item.Rights } else { "" }
                                        if ($identity -or $rights) { $formattedItems += "<strong>$identity</strong> - access:`n $rights" }
                                    }
                                }
                                else {
                                    # Single object
                                    $identity = if ($jsonValue.Identity) { $jsonValue.Identity } else { "" }
                                    $rights = if ($jsonValue.Rights) { $jsonValue.Rights } else { "" }
                                    if ($identity -or $rights) {
                                        $formattedItems += "<strong>$identity</strong> - access:`n $rights"
                                    }
                                }

                                # Display all items in the Value column, separated by <br/>
                                if ($formattedItems) {
                                    $displayValue = $formattedItems -join "<br/>"
                                    "<tr><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Key)</td><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$displayValue</td></tr>" | Add-Content -Path $HTML
                                }
                            }
                        }
                        catch { "<tr><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Key)</td><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Value)</td></tr>" | Add-Content -Path $HTML }
                    }
                    # Special handling for Protection level - extract DisplayName from JSON
                    elseif ($Setting.Key -eq "Protection level" -and $Setting.Value) {
                        try {
                            $jsonValue = $Setting.Value | ConvertFrom-Json
                            if ($jsonValue -and $jsonValue.DisplayName) { "<tr><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Key)</td><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($jsonValue.DisplayName)</td></tr>" | Add-Content -Path $HTML }
                            else { "<tr><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Key)</td><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Value)</td></tr>" | Add-Content -Path $HTML }
                        }
                        catch { "<tr><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Key)</td><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Value)</td></tr>" | Add-Content -Path $HTML }
                    }
                    else { "<tr><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Key)</td><td style='border: 1px solid #ddd; padding: 8px; word-break: break-word;'>$($Setting.Value)</td></tr>" | Add-Content -Path $HTML }
                }
                "</table>" | Add-Content -Path $HTML
            }
        }

        "</div>" | Add-Content -Path $HTML
    }

    ##########################################################################################################################################################
    # Render labels in overall priority order

    foreach ($TopLabel in $TopLevelLabels) {

        if ($TopLabel.IsLabelGroup -eq $true) {
            "<details class='label-group'><summary class='label-group-header'>Label Group - $($TopLabel.DisplayName)&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;Priority: $($TopLabel.Priority)</summary><div class='label-group-content'>" | Add-Content -Path $HTML
            $SubLabels = @(
                $Labels |
                Where-Object {
                    $_.IsLabelGroup -ne $true -and
                    ([string]$_.ParentId).Trim() -eq ([string]$TopLabel.ImmutableId).Trim()
                } |
                Sort-Object Priority
            )

            if ($SubLabels.Count -eq 0) { "<p>There are no labels in this label group.</p>" | Add-Content -Path $HTML }
            else {
                foreach ($Label in $SubLabels) { & $WriteLabel -Label $Label }
            }

            '</div></details>' | Add-Content -Path $HTML
        }
        else { & $WriteLabel -Label $TopLabel }
    }

    # Close the Sensitivity Labels section-content div
    "</div>" | Add-Content -Path $HTML
}


##########################################################################################################################################################


function labelPolicies {

    Write-Host "Gathering Sensitivity Label Policies"
    #Sensitivity Label Policies title
    '<h2>Label Policies</h2><div class="section-content">' | Add-Content -Path $HTML

    ##########################################################################################################################################################

    #Get Sensitivity Label publishing policy
    $LabelPolicies = Get-LabelPolicy -WarningAction SilentlyContinue
    Write-Host "Found $($LabelPolicies.Count) sensitivity label policies"

    # Hash tables for translating label policy settings (similar to labels function)
    $policyActionLabels = @{
        "powerbimandatory"              = "Require users to apply a label to their Fabric and Power BI content"
        "requiredowngradejustification" = "Users must provide a justification to remove a label or lower its classification"
        "mandatory"                     = "Require users to apply a label to their emails and documents"
        "customurl"                     = "Provide users with a link to a custom help page"
        "outlookdefaultlabel"           = "Default label for emails"
        "disablemandatoryinoutlook"     = "Disable mandatory labeling for emails"
        "defaultlabelid"                = "Default label for documents"
        "attachmentaction"              = "Email inherits highest priority label from attachments"
        "inheritlabelfromcontent"       = "Inherit label from files shared to meetings"
        "applylabeltoartifacts"         = "Apply meeting label to artifacts"
        "teamworkmandatory"             = "Require users to apply a label to their meetings and calendar events"
        "siteandgroupmandatory"         = "Require users to apply a label to their groups or sites"
        "siteandgroupdefaultlabelid"    = "Default label for sites and groups"
        "powerbidefaultlabelid"         = "Default label for Fabric and Power BI content"
        "teamworkdefaultlabelid"        = "Default label for meeting and calendar events"
    }
    $policyValueLabels = @{
        "true"        = "True"
        "false"       = "False"
        "automatic"   = "Automatic"
        "recommended" = "Recommended"
    }

    #Foreach Label Policy
    foreach ($policy in $LabelPolicies) {
        "<h3>$($policy.Name) - Priority: $($policy.Priority)</h3><div class='h3-content'>" | Add-content $HTML
        if ([string]::IsNullOrEmpty($policy.CreatedBy)) { $policy.CreatedBy = "Microsoft Corporation" }        
        "<div style='background-color: #f8f9fa; padding: 10px; border-radius: 4px; margin: 0 0 8px 0; color: #212529;'><strong>Created by:</strong> $($policy.CreatedBy)<br><strong>Last modified by:</strong> $($policy.LastModifiedBy)<br><strong>Last modified:</strong> $(if ($policy.WhenChanged) { $policy.WhenChanged.ToString('yyyy-MM-dd') } else { '' })</div>" | Add-content $HTML


        #Desc/info - always use label-like styling
        if ([string]::IsNullOrEmpty($policy.Comment)) { '<div style="background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;"><strong>Description:</strong> No description provided</div>' | Add-content $HTML }
        else { "<div style='background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Description:</strong> $($policy.Comment)</div>" | Add-content $HTML }

        #Published Labels per policy
        '<div style="background-color: #e8f5e8; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;">
            <h4 style="margin: 0;">Published Labels:</h4>' | Add-Content $HTML
        $labelNames = foreach ($labelItem in $policy.Labels) {
            $isGroup = $false
            $guid = $labelItem
            if ($labelItem.EndsWith("Group")) {
                $isGroup = $true
                $guid = $labelItem.Substring(0, $labelItem.Length - 5)
            }
            if ([string]::IsNullOrEmpty($guid)) { continue }
            $label = $script:Labels | Where-Object { $_.ImmutableId -eq $guid } | Select-Object -First 1
            if ($label) {
                if ($isGroup) { "$($label.DisplayName) (Group/Header label)" }
                else { $label.DisplayName }
            }
            else { $labelItem }
        }
        $formattedLabels = ($labelNames -join "<br>")
        '<p style="margin: 0;">' + $formattedLabels + '</p>' | Add-Content $HTML
        '</div>' | Add-Content $HTML

        ###########AdminUnits per policy
        #$AdminUnits = $policy.PolicyConstraints | ForEach-Object { $_ | ConvertFrom-Json }
        #"<h4>Administrative units: </h4>" | Add-Content $HTML

        #Published To per policy
        "<h4>Published to:</h4>" | Add-Content $HTML
        "<table><tr><th>Container</th><th>Groups/Users</th></tr>" | Add-Content $HTML

        #Exchange
        $LabelPolicyUsers = ($policy.ExchangeLocation -join "<br>")
        if (-not([string]::IsNullOrEmpty($policy.ExchangeLocation))) { "<tr><td>Users</td><td>$($LabelPolicyUsers)</td></tr>" | Add-Content $HTML }
        if (-not([string]::IsNullOrEmpty($policy.ExchangeLocationException))) { "<tr><td>User Exceptions</td><td>$($policy.ExchangeLocationException)</td></tr>" | Add-Content $HTML }
        #M365 Groups
        $LabelPolicyGroups = ($policy.ModernGroupLocation -join "<br>")
        if (-not([string]::IsNullOrEmpty($policy.ModernGroupLocation))) { "<tr><td>Groups</td><td>$($LabelPolicyGroups)</td></tr>" | Add-Content $HTML }
        if (-not([string]::IsNullOrEmpty($policy.ModernGroupLocationException))) { "<tr><td>Group Exceptions</td><td>$($policy.ModernGroupLocationException)</td></tr>" | Add-Content $HTML }
        "</table>" | Add-Content $HTML

        #Parse the XML
        $xml = [xml]$policy.PolicySettingsBlob
        "<h4>Policy Settings:</h4>" | Add-Content $HTML
        "<Table><tr><th>Configuration</th><th>Value</th></tr>" | Add-Content $HTML
        #Foreach setting in the policy
        $labelIdKeys = @("defaultlabelid", "teamworkdefaultlabelid", "outlookdefaultlabel", "siteandgroupdefaultlabelid", "powerbidefaultlabelid")
        foreach ($setting in $xml.settings.setting) {
            $rawKey = $setting.key
            $rawValue = $setting.value

            # Translate the key if possible
            $displayKey = $policyActionLabels[$rawKey]
            if (-not $displayKey) { $displayKey = $rawKey }

            # Check if this is a label ID key
            if ($labelIdKeys -contains $rawKey) {
                $label = $script:Labels | Where-Object { $_.ImmutableId -eq $rawValue } | Select-Object -First 1
                if ($label) { $displayValue = $label.DisplayName }
                else {
                    $displayValue = $policyValueLabels[$rawValue]
                    if (-not $displayValue) { $displayValue = $rawValue }
                }
            }
            else {
                $displayValue = $policyValueLabels[$rawValue]
                if (-not $displayValue) { $displayValue = $rawValue }
            }

            "<tr><td>$displayKey</td><td>$displayValue</td></tr>" | Add-Content $HTML
        }
        "</table>" | Add-Content $HTML
        #Close section-content div
        "</div>" | Add-Content $HTML
    }

    #Close section-content div
    "</div>" | Add-Content -Path $HTML



}


#################################################################################################################################


function AutoLabelingPolicy {

    Write-Host "Gathering Auto-Labeling Policies"
    '<h2>Auto-Labeling Policies</h2><div class="section-content">' | Add-Content -Path $HTML

    #################################################################################################################################


    $AutoLabeling = Get-AutoSensitivityLabelPolicy

    # Sort policies by Priority
    $AutoLabeling = $AutoLabeling | Sort-Object Priority

    foreach ($AutoPol in $AutoLabeling) {
        if ($AutoPol.RemoveLabel -eq "True") { "<h3>$($AutoPol.Name) - Priority: $($AutoPol.Priority) - Remove labels automatically</h3>" | Add-Content -Path $HTML }
        else { "<h3>$($AutoPol.Name) - Priority: $($AutoPol.Priority) - Apply labels automatically</h3>" | Add-Content -Path $HTML }
        "<div class='h3-content'>" | Add-Content -Path $HTML
        "<div style='background-color: #f8f9fa; padding: 10px; border-radius: 4px; margin: 0 0 8px 0; color: #212529;'><strong>Created by:</strong> $($AutoPol.CreatedBy)<br><strong>Last modified by:</strong> $($AutoPol.LastModifiedBy)<br><strong>Last modified:</strong> $(if ($AutoPol.WhenChanged) { $AutoPol.WhenChanged.ToString('yyyy-MM-dd') } else { '' })</div>" | Add-Content -Path $HTML

        #Desc/info - always use label-like styling
        if ([string]::IsNullOrEmpty($AutoPol.Comment)) { '<div style="background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;"><strong>Description:</strong> No description provided</div>' | Add-content $HTML }
        else { "<div style='background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Description: </strong>$($AutoPol.Comment)</div>" | Add-content $HTML }

        "<div style='background-color: #e8f5e8; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Mode:</strong> $($polMode[$AutoPol.Mode])</div>" | Add-content $HTML
        
        if ($AutoPol.RemoveLabel -ne "True") {
            $labelName = $($AutoPol.ApplySensitivityLabel)
            $labelInfo = $script:Labels | Where-Object { $_.ImmutableId -eq $labelName } | Select-Object -First 1
            if ($labelInfo) {
                $displayLabelName = $labelInfo.DisplayName
            }
            else {
                $displayLabelName = $labelName
            }
            if ($AutoPol.isDefaultSPODocLibraryLabelInUse -eq "True") { "<div style='background-color: #e9ecef; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Label assigned:</strong> Same as the default sensitivity label of the SharePoint document library.</div>" | Add-content $HTML }
            else { "<div style='background-color: #e9ecef; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Label assigned:</strong> $displayLabelName</div>" | Add-content $HTML }
        }

        # Add the location table
        "<table><tr><th>Location</th><th>Assigned to:</th></tr>" | Add-Content $HTML

        # Exchange Location
        $exchangeSender = if ($AutoPol.ExchangeSender) { $AutoPol.ExchangeSender } else { @() }
        $exchangeLocation = if ($AutoPol.ExchangeLocation) { $AutoPol.ExchangeLocation } else { @() }
        $exchangeSenderException = if ($AutoPol.ExchangeSenderException) { $AutoPol.ExchangeSenderException } else { @() }
        $exchangeSenderMemberOfException = if ($AutoPol.ExchangeSenderMemberOfException) { $AutoPol.ExchangeSenderMemberOfException } else { @() }
        $exchangeLocationException = if ($AutoPol.ExchangeLocationException) { $AutoPol.ExchangeLocationException } else { @() }
        $exchangeAdaptiveScopesException = if ($AutoPol.ExchangeAdaptiveScopesException) { $AutoPol.ExchangeAdaptiveScopesException } else { @() }

        # Determine Exchange Location value: ExchangeSender if populated, else ExchangeLocation
        $exchangeValue = @()
        if ($exchangeSender.Count -gt 0) { $exchangeValue = $exchangeSender }
        elseif ($exchangeLocation.Count -gt 0) { $exchangeValue = $exchangeLocation }

        if ($exchangeValue.Count -gt 0) {
            "<tr><td>Exchange Location</td><td>$($exchangeValue -join ', ')</td></tr>" | Add-Content $HTML
            # Show Exchange Sender Exception if it has value
            if ($exchangeSenderException.Count -gt 0) {
                # Format the sender exception details for display
                $senderExceptionDetails = $exchangeSenderException | ForEach-Object {
                    if ($_.PSObject.TypeNames -contains 'System.String') {
                        # Check if the string looks like JSON
                        $trimmed = $_.TrimStart()
                        if ($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
                            try {
                                $jsonObj = $_ | ConvertFrom-Json
                                if ($jsonObj.DisplayName) { "$($jsonObj.DisplayName) (User)" }
                                elseif ($jsonObj.PrimarySmtpAddress) { "$($jsonObj.PrimarySmtpAddress) (User)" }
                                else { "$_ (User)" }
                            }
                            catch { "$_ (User)" }
                        }
                        else { "$_ (User)" }
                    }
                    else {
                        # Regular object access
                        if ($_.DisplayName) { "$($_.DisplayName) (User)" }
                        elseif ($_.PrimarySmtpAddress) { "$($_.PrimarySmtpAddress) (User)" }
                        else { "$_ (User)" }
                    }
                }
                "<tr><td>Exchange Sender Exception</td><td>$($senderExceptionDetails -join '<br>')</td></tr>" | Add-Content $HTML
            }
            # Show Exchange Sender Member Of Exception if it has value
            if ($exchangeSenderMemberOfException.Count -gt 0) {
                # Format the sender member of exception details for display
                $senderMemberOfExceptionDetails = $exchangeSenderMemberOfException | ForEach-Object {
                    if ($_.PSObject.TypeNames -contains 'System.String') {
                        # Check if the string looks like JSON
                        $trimmed = $_.TrimStart()
                        if ($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
                            try {
                                $jsonObj = $_ | ConvertFrom-Json
                                if ($jsonObj.DisplayName) { "$($jsonObj.DisplayName) (Group)" }
                                elseif ($jsonObj.PrimarySmtpAddress) { "$($jsonObj.PrimarySmtpAddress) (Group)" }
                                else { "$_ (Group)" }
                            }
                            catch { "$_ (Group)" }
                        }
                        else { "$_ (Group)" }
                    }
                    else {
                        # Regular object access
                        if ($_.DisplayName) { "$($_.DisplayName) (Group)" }
                        elseif ($_.PrimarySmtpAddress) { "$($_.PrimarySmtpAddress) (Group)" }
                        else { "$_ (Group)" }
                    }
                }
                "<tr><td>Exchange Sender Member Of Exception</td><td>$($senderMemberOfExceptionDetails -join '<br>')</td></tr>" | Add-Content $HTML
            }
            # Show Exchange Location Exception if it has value
            if ($exchangeLocationException.Count -gt 0) {
                # Format the location exception details for display
                $locationExceptionDetails = $exchangeLocationException | ForEach-Object {
                    if ($_.PSObject.TypeNames -contains 'System.String') {
                        # Check if the string looks like JSON
                        $trimmed = $_.TrimStart()
                        if ($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
                            try {
                                $jsonObj = $_ | ConvertFrom-Json
                                if ($jsonObj.PrimarySmtpAddress) { "$($jsonObj.PrimarySmtpAddress) (User)" }
                                elseif ($jsonObj.DisplayName) { "$($jsonObj.DisplayName) (User)" }
                                else { "$_ (User)" }
                            }
                            catch { "$_ (User)" }
                        }
                        else { "$_ (User)" }
                    }
                    else {
                        if ($_.PrimarySmtpAddress) { "$($_.PrimarySmtpAddress) (User)" }
                        elseif ($_.DisplayName) { "$($_.DisplayName) (User)" }
                        else { "$_ (User)" }
                    }
                }
                "<tr><td>Exchange Location Exception</td><td>$($locationExceptionDetails -join '<br>')</td></tr>" | Add-Content $HTML
            }
            # Show Exchange Adaptive Scopes Exception if it has value
            if ($exchangeAdaptiveScopesException.Count -gt 0) {
                # Format the adaptive scopes exception details for display
                $adaptiveScopesExceptionDetails = $exchangeAdaptiveScopesException | ForEach-Object {
                    # Try to parse as JSON if it's a string
                    if ($_.PSObject.TypeNames -contains 'System.String') {
                        try {
                            $jsonObj = $_ | ConvertFrom-Json
                            # For adaptive scopes, we likely want to show the raw value or a name field
                            if ($jsonObj -is [System.String]) { "$jsonObj (Adaptive Scope)" }
                            elseif ($jsonObj.DisplayName) { "$($jsonObj.DisplayName) (Adaptive Scope)" }
                            elseif ($jsonObj.Name) { "$($jsonObj.Name) (Adaptive Scope)" }
                            else { "$_ (Adaptive Scope)" }
                        }
                        catch { "$_ (Adaptive Scope)" }
                    }
                    else {
                        # Regular object access
                        if ($_) { "$_ (Adaptive Scope)" }
                        else { "$_ (Adaptive Scope)" }
                    }
                }
                "<tr><td>Exchange Adaptive Scopes Exception</td><td>$($adaptiveScopesExceptionDetails -join '<br>')</td></tr>" | Add-Content $HTML
            }
            # If no specific exceptions have values, check for N/A logic
            else {
                # Check if we should show N/A (when primary is populated but all exceptions are empty)
                if (($exchangeSenderException.Count -eq 0) -and ($exchangeSenderMemberOfException.Count -eq 0) -and ($exchangeLocationException.Count -eq 0) -and ($exchangeAdaptiveScopesException.Count -eq 0)) {
                    "<tr><td>Exchange Location Exception</td><td>N/A</td></tr>" | Add-Content $HTML
                }
            }
        }

        # SharePoint Location
        $sharePointLocation = if ($AutoPol.SharePointLocation) { $AutoPol.SharePointLocation } else { @() }
        $sharePointLocationException = if ($AutoPol.SharePointLocationException) { $AutoPol.SharePointLocationException } else { @() }
        $sharePointAdaptiveScopesException = if ($AutoPol.SharePointAdaptiveScopesException) { $AutoPol.SharePointAdaptiveScopesException } else { @() }

        if ($sharePointLocation.Count -gt 0) {
            "<tr><td>SharePoint Location</td><td>$($sharePointLocation -join ', ')</td></tr>" | Add-Content $HTML

            # Show SharePoint Location Exception (site) if it has value
            if ($sharePointLocationException.Count -gt 0) { "<tr><td>SharePoint Location Exception (Site)</td><td>$($sharePointLocationException -join '<br>')</td></tr>" | Add-Content $HTML }
            if ($sharePointAdaptiveScopesException.Count -gt 0) { "<tr><td>SharePoint Location Exception (Adaptive Scope)</td><td>$($sharePointAdaptiveScopesException -join '<br>')</td></tr>" | Add-Content $HTML }
            else {
                if (($sharePointLocationException.Count -eq 0) -and ($sharePointAdaptiveScopesException.Count -eq 0)) { "<tr><td>SharePoint Location Exception</td><td>N/A</td></tr>" | Add-Content $HTML }
            }
        }

        # OneDrive Location
        $oneDriveLocation = if ($AutoPol.OneDriveLocation) { $AutoPol.OneDriveLocation } else { @() }
        $exceptIfOneDriveSharedBy = if ($AutoPol.ExceptIfOneDriveSharedBy) { $AutoPol.ExceptIfOneDriveSharedBy } else { @() }
        $exceptIfOneDriveSharedByMemberOf = if ($AutoPol.ExceptIfOneDriveSharedByMemberOf) { $AutoPol.ExceptIfOneDriveSharedByMemberOf } else { @() }

        if ($oneDriveLocation.Count -gt 0) {
            "<tr><td>OneDrive Location</td><td>$($oneDriveLocation -join ', ')</td></tr>" | Add-Content $HTML

            # Show OneDrive Shared By Exception (user) if it has value
            if ($exceptIfOneDriveSharedBy.Count -gt 0) {
                # Format the user details for display
                $userDetails = $exceptIfOneDriveSharedBy | ForEach-Object {
                    # Try to parse as JSON if it's a string
                    if ($_.PSObject.TypeNames -contains 'System.String') {
                        try {
                            $jsonObj = $_ | ConvertFrom-Json
                            if ($jsonObj.PrimarySmtpAddress) { "$($jsonObj.PrimarySmtpAddress) (User)" }
                            elseif ($jsonObj.DisplayName) { "$($jsonObj.DisplayName) (User)" }
                            else { "$_ (User)" }
                        }
                        catch { "$_ (User)" }
                    }
                    else {
                        if ($_.PrimarySmtpAddress) { "$($_.PrimarySmtpAddress) (User)" }
                        elseif ($_.DisplayName) { "$($_.DisplayName) (User)" }
                        else { "$_ (User)" }
                    }
                }
                "<tr><td>OneDrive Shared By Exception (User)</td><td>$($userDetails -join '<br>')</td></tr>" | Add-Content $HTML
            }
            # Show OneDrive Shared By Member Of Exception (group) if it has value
            if ($exceptIfOneDriveSharedByMemberOf.Count -gt 0) {
                # Format the group details for display
                $groupDetails = $exceptIfOneDriveSharedByMemberOf | ForEach-Object {
                    # Try to parse as JSON if it's a string
                    if ($_.PSObject.TypeNames -contains 'System.String') {
                        try {
                            $jsonObj = $_ | ConvertFrom-Json
                            if ($jsonObj.DisplayName) { "$($jsonObj.DisplayName) (Group)" }
                            elseif ($jsonObj.PrimarySmtpAddress) { "$($jsonObj.PrimarySmtpAddress) (Group)" }
                            else { "$_ (Group)" }
                        }
                        catch { "$_ (Group)" }
                    }
                    else {
                        # Regular object access
                        if ($_.DisplayName) { "$($_.DisplayName) (Group)" }
                        elseif ($_.PrimarySmtpAddress) { "$($_.PrimarySmtpAddress) (Group)" }
                        else { "$_ (Group)" }
                    }
                }
                "<tr><td>OneDrive Shared By Member Of Exception (Group)</td><td>$($groupDetails -join '<br>')</td></tr>" | Add-Content $HTML
            }
            # If no specific exceptions have values, check for N/A logic
            else {
                # Check if we should show N/A (when primary is populated but all exceptions are empty)
                if (($exceptIfOneDriveSharedBy.Count -eq 0) -and ($exceptIfOneDriveSharedByMemberOf.Count -eq 0)) {
                    "<tr><td>OneDrive Location Exception</td><td>N/A</td></tr>" | Add-Content $HTML
                }
            }
        }

        "</table>" | Add-Content $HTML

        if ($AutoPol.isDefaultSPODocLibraryLabelInUse -ne "True") {
            $AutoPolRule = Get-AutoSensitivityLabelRule -Policy $AutoPol
            foreach ($AuRule in $AutoPolRule) {
                function Get-ConditionDisplayName {
                    param ([string]$ConditionName)

                    $allConditionsCondition = @{
                        "ContentContainsSensitiveInformation" = "Content Contains:"
                        "AccessScope"                         = "Content is shared:"
                        "SenderIPRanges"                      = "Sender IP address is:"
                        "RecipientDomainIs"                   = "Recipient domain is:"
                        "ContentExtensionMatchesWords"        = "Attachment extension is:"
                        "DocumentIsPasswordProtected"         = "Attachment is password protected"
                        "DocumentIsUnsupported"               = "Any email attachment's content could not be scanned"
                        "ProcessingLimitExceeded"             = "Any email attachment's content didn't complete scanning"
                        "HeaderMatchesPatterns"               = "Header matches patterns:"
                        "SubjectMatchesPatterns"              = "Subject matches patterns:"
                        "AnyOfRecipientAddressContainsWords"  = "Recipient address contains words:"
                        "FromAddressMatchesPatterns"          = "Sender address matches pattern"
                        "SenderDomainIs"                      = "Sender domain is:"
                        "SentToMemberOf"                      = "Recipient is a member of:"
                        "From"                                = "Sender is:"
                        "DocumentNameMatchesWords"            = "Attachment name matches words or phrases:"
                        "ContentPropertyContainsWords"        = "Attachment property is:"
                        "DocumentSizeOver"                    = "Attachment size equals or is greater than:"
                        "SentTo"                              = "Recipient is:"

                        "DocumentCreatedBy"                   = "Document created by:"
                        "DocumentCreatedOnOrAfter"            = "Document creation date is on or after"
                        "DocumentCreatedOnOrBefore"           = "Document creation date is on or before"
                        "DocumentLastModifiedOnOrAfter"       = "Document modification date is on or after"
                        "DocumentLastModifiedOnOrBefore"      = "Document modification date is on or before"
                    }

                    if ($allConditionsCondition.ContainsKey($ConditionName)) { return $allConditionsCondition[$ConditionName] }
                    return $ConditionName
                }

                function Get-ValueDisplayName {
                    param (
                        [AllowNull()]
                        $Value
                    )

                    $allConditionsValue = @{
                        "NotInOrganization" = "With people outside my organization"
                        "InOrganization"    = "Only with people inside my organization"
                    }

                    if (
                        $null -ne $Value -and
                        $allConditionsValue.ContainsKey([string]$Value)
                    ) { return $allConditionsValue[[string]$Value] }
                    return $Value
                }

                function Write-ConditionValue {
                    param (
                        $Condition,
                        [string]$HTML
                    )

                    if ($Condition.ConditionName -eq "ContentContainsSensitiveInformation") {
                        foreach ($valueItem in @($Condition.Value)) {
                            $groupCount = 0
                            foreach ($sitGroup in @($valueItem.Groups)) {
                                $groupCount++
                                "<div class='sit-group'>" | Add-Content -Path $HTML
                                "<strong>Group name:</strong> $($sitGroup.Name)" | Add-Content -Path $HTML
                                if ($($sitGroup.Operator) -eq "And") { "<strong> Group operator:</strong> All of these<br>" | Add-Content -Path $HTML }
                                else { "<strong> Group operator:</strong> Any of these<br>" | Add-Content -Path $HTML }
                                $regularTypes = @()
                                $trainableTypes = @()
                                foreach ($sensitiveType in @($sitGroup.Sensitivetypes)) {
                                    if ($sensitiveType.Classifiertype -eq "MLModel") {
                                        $trainableTypes += $sensitiveType
                                    }
                                    else {
                                        $regularTypes += $sensitiveType
                                    }
                                }
                                $meaningfulRegularTypes = $regularTypes | Where-Object {
                                    $_.Confidencelevel -and $_.Confidencelevel.Trim() -ne ''
                                }
                                if ($meaningfulRegularTypes.Count -gt 0) {
                                    "<div class='sit-item'>" | Add-Content -Path $HTML
                                    "<strong>&nbsp;&nbsp;&nbsp;Sensitive Information Types:</strong><br>" | Add-Content -Path $HTML
                                    "</div>" | Add-Content -Path $HTML
                                    foreach ($sensitiveType in $meaningfulRegularTypes) {
                                        "<div class='sit-item'>" | Add-Content -Path $HTML
                                        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$($sensitiveType.Name) --- $($sensitiveType.Confidencelevel) --- $($sensitiveType.Mincount) to " | Add-Content -Path $HTML
                                        if ($($sensitiveType.Maxcount) -eq "-1") { "Any" | Add-Content -Path $HTML }
                                        else { "$($sensitiveType.Maxcount)<br>" | Add-Content -Path $HTML }
                                        "</div>" | Add-Content -Path $HTML
                                    }
                                }
                                if ($trainableTypes.Count -gt 0) {
                                    "<div class='sit-item'>" | Add-Content -Path $HTML
                                    "<strong>&nbsp;&nbsp;&nbsp;Trainable Classifiers:</strong><br>" | Add-Content -Path $HTML
                                    "</div>" | Add-Content -Path $HTML
                                    foreach ($sensitiveType in $trainableTypes) {
                                        "<div class='sit-item'>" | Add-Content -Path $HTML
                                        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$($sensitiveType.Name)<br>" | Add-Content -Path $HTML
                                        "</div>" | Add-Content -Path $HTML
                                    }
                                }
                                # Handle Sensitivity Labels
                                $sensitivityLabels = @()
                                if ($sitGroup.Labels) {
                                    foreach ($label in @($sitGroup.Labels)) {
                                        if ($label.Type -eq "Sensitivity") { $sensitivityLabels += $label }
                                    }
                                }
                                if ($sensitivityLabels.Count -gt 0) {
                                    "<div class='sit-item'>" | Add-Content -Path $HTML
                                    "<strong>&nbsp;&nbsp;&nbsp;Sensitivity Labels:</strong><br>" | Add-Content -Path $HTML
                                    "</div>" | Add-Content -Path $HTML
                                    foreach ($label in $sensitivityLabels) {
                                        # Find the label in $script:Labels by ImmutableId (matching against label.Id or label.Name from JSON)
                                        $labelInfo = $script:Labels | Where-Object { $_.ImmutableId -eq $label.Id -or $_.ImmutableId -eq $label.Name } | Select-Object -First 1
                                        if ($labelInfo) {
                                            "<div class='sit-item'>" | Add-Content -Path $HTML
                                            "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$($labelInfo.DisplayName)<br>" | Add-Content -Path $HTML
                                            "</div>" | Add-Content -Path $HTML
                                        }
                                        else {
                                            # Fallback to showing the ID if label not found
                                            "<div class='sit-item'>" | Add-Content -Path $HTML
                                            "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$($label.Id)<br>" | Add-Content -Path $HTML
                                            "</div>" | Add-Content -Path $HTML
                                        }
                                    }
                                }
                                "</div>" | Add-Content -Path $HTML
                                # Show operator BETWEEN groups (but not after the last group)
                                if ($groupCount -lt $valueItem.Groups.Count) {
                                    "<div class='sit-group-operator'>$($valueItem.Operator)</div>" | Add-Content -Path $HTML
                                }
                            }
                        }

                        return
                    }
                    if ($null -ne $Condition.Value) {
                        if ($Condition.ConditionName.Trim() -eq "HeaderMatchesPatterns") {
                            # Handle HeaderMatchesPatterns specially: show key and sub-values
                            try {
                                if ($Condition.Value) {
                                    $propertyNames = $Condition.Value.PSObject.Properties.Name
                                    if ($propertyNames) {
                                        foreach ($key in $propertyNames) {
                                            "<strong>&nbsp;&nbsp;&nbsp;Header:</strong> $key<br>" | Add-Content -Path $HTML
                                            "<strong>&nbsp;&nbsp;&nbsp;Values:</strong><br>" | Add-Content -Path $HTML
                                            $subValues = @($Condition.Value.$key)
                                            if ($subValues.Count -gt 0) {
                                                foreach ($subValue in $subValues) {
                                                    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$subValue<br>" | Add-Content -Path $HTML
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch {
                                # Fallback to generic handling if special processing fails
                                $displayValues = foreach ($valueItem in @($Condition.Value)) { Get-ValueDisplayName -Value $valueItem }
                                "&nbsp;&nbsp;&nbsp;$($displayValues -join '<br>&nbsp;&nbsp;&nbsp;')<br>" | Add-Content -Path $HTML
                            }
                        }
                        else {
                            $displayValues = foreach ($valueItem in @($Condition.Value)) { Get-ValueDisplayName -Value $valueItem }
                            # Join with <br> instead of comma for new lines
                            "&nbsp;&nbsp;&nbsp;$($displayValues -join '<br>&nbsp;&nbsp;&nbsp;')<br>" | Add-Content -Path $HTML
                        }
                    }
                    if ($Condition.GlobalListIds) { "GlobalListIds: $(@($Condition.GlobalListIds) -join ', ')<br>" | Add-Content -Path $HTML }
                }

                function Write-ConditionGroup {
                    param (
                        $Condition,
                        [string]$HTML,
                        [int]$Depth = 0,
                        [string]$GroupName = "Group"
                    )


                    "<div class='condition-group group-depth-$Depth'>" | Add-Content -Path $HTML
                    "<div class='condition-group-heading'>" | Add-Content -Path $HTML
                    "<strong>$GroupName</strong> " | Add-Content -Path $HTML
                    "</div>" | Add-Content -Path $HTML

                    $nestedGroupNumber = 0
                    foreach ($Item in @($Condition.SubConditions)) {

                        if ($null -ne $Item.ConditionName) {
                            $conditionDisplayName = Get-ConditionDisplayName -ConditionName $Item.ConditionName
                            "<div class='condition-item'>" | Add-Content -Path $HTML
                            "<strong>$conditionDisplayName</strong><br>" | Add-Content -Path $HTML
                            Write-ConditionValue -Condition $Item -HTML $HTML
                            "</div>" | Add-Content -Path $HTML
                        }
                        elseif ($null -ne $Item.SubConditions) {
                            $nestedGroupNumber++
                            Write-ConditionGroup -Condition $Item -HTML $HTML -Depth ($Depth + 1) -GroupName "Group $nestedGroupNumber"
                        }
                    }

                    "</div>" | Add-Content -Path $HTML
                }

                $AutoRule = Get-AutoSensitivityLabelRule -Identity $AuRule.Identity
                $test = $AutoRule.AdvancedRule | ConvertFrom-Json
                "<div style='background-color: #e8f5e8; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'>" | Add-Content -Path $HTML
                "<strong>$($AutoRule.DisplayName):</strong>" | Add-Content -Path $HTML
                "</div>" | Add-Content -Path $HTML
                "<div class='advanced-rules-container'>" | Add-Content -Path $HTML
                function GroupBox {
                    param (
                        $Condition,
                        [string]$HTML,
                        [int]$Index,
                        [int]$Depth
                    )
                    "<div class='condition-group group-depth-$Depth'>" | Add-Content -Path $HTML
                    "<div class='condition-group-heading'>" | Add-Content -Path $HTML
                    "<strong>Group $Index</strong>" | Add-Content -Path $HTML
                    if ($Condition.Operator -eq "Not") { " <span class='condition-not-operator'>(Not)</span>" | Add-Content -Path $HTML }
                    "</div>" | Add-Content -Path $HTML
                    if ($Condition.SubConditions) {
                        $nestedGroupNumber = 0
                        $itemCount = 0
                        foreach ($Item in @($Condition.SubConditions)) {
                            $itemCount++
                            if ($null -ne $Item.ConditionName) {
                                $conditionDisplayName = Get-ConditionDisplayName -ConditionName $Item.ConditionName
                                "<div class='condition-item'>" | Add-Content -Path $HTML
                                "<strong>$conditionDisplayName</strong><br>" | Add-Content -Path $HTML
                                Write-ConditionValue -Condition $Item -HTML $HTML
                                "</div>" | Add-Content -Path $HTML
                            }
                            elseif ($null -ne $Item.SubConditions) {
                                $nestedGroupNumber++
                                GroupBox -Condition $Item -HTML $HTML -Index $nestedGroupNumber -Depth ($Depth + 1)
                            }
                            # Show operator BETWEEN subconditions (but not after the last item)
                            if ($itemCount -lt $Condition.SubConditions.Count) {
                                "<div class='condition-operator-between-left'>$($Condition.Operator)</div>" | Add-Content -Path $HTML
                            }
                        }
                    }
                    else {
                        # Render the condition itself as a condition-item
                        $conditionDisplayName = Get-ConditionDisplayName -ConditionName $Condition.ConditionName
                        "<div class='condition-item'>" | Add-Content -Path $HTML
                        "<strong>$conditionDisplayName</strong><br>" | Add-Content -Path $HTML
                        Write-ConditionValue -Condition $Condition -HTML $HTML
                        "</div>" | Add-Content -Path $HTML
                    }
                    "</div>" | Add-Content -Path $HTML
                }
                $groupIndex = 0
                $conditionOperator = if ($test.Condition.Operator) { $test.Condition.Operator } else { "And" }
                foreach ($subCondition in @($test.Condition.SubConditions)) {
                    $groupIndex++
                    GroupBox -Condition $subCondition -HTML $HTML -Index $groupIndex -Depth 0
                    # Add operator between groups (except after the last group)
                    if ($groupIndex -lt $test.Condition.SubConditions.Count) {
                        "<div class='condition-operator-between'>$conditionOperator</div>" | Add-Content -Path $HTML
                    }
                }
                "</div>" | Add-Content -Path $HTML
            }
        }

        $autoPolSettings = @{}
        $autoPolSettings["ApplySensitivityLabelOverwriteWorkloads"] = $AutoPol.ApplySensitivityLabelOverwriteWorkloads
        $autoPolSettings["ExternalMailRightsManagementOwner"] = $AutoPol.ExternalMailRightsManagementOwner


        $autoPolSettingsValue = @{
            "Exchange, SharePoint, OneDriveForBusiness" = "All locations"
            "Exchange"                                  = "Emails only"
        }
        if ($autoPolSettings.Count -ne 0) {
            "<div style='background-color: #e9ecef; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><Strong>Additional label settings</Strong><br>" | Add-Content $HTML 
            if ($autoPolSettings.ApplySensitivityLabelOverwriteWorkloads) { "Replace existing labels if their priority is the same or lower for: <Strong>$($autoPolSettingsValue[$autoPolSettings['ApplySensitivityLabelOverwriteWorkloads']])</Strong><br>" | Add-Content $HTML }
            if ($autoPolSettings.ExternalMailRightsManagementOwner) { "Apply encryption to email received from outside your organization. Rights Management owner - <Strong>$($autoPolSettings.ExternalMailRightsManagementOwner)</Strong>" | Add-Content $HTML }
            "</div>" | Add-Content $HTML 
        }


        #Close section-content div
        "</div>" | Add-Content $HTML
    }
    "</div>" | Add-Content -Path $HTML
}


#################################################################################################################################


function DLPPolicies {

    Write-Host "Gathering DLP Policies"
    #DLP
    '<h2>Data Loss Prevention</h2><div class="section-content">' | Add-Content -Path $HTML

    #################################################################################################################################
    #Get-DlpComplianceRule
    $DLPPolicy = Get-DlpCompliancePolicy
    Write-Host "Found $($DLPPolicy.Count) DLP policies"

    foreach ($DLPPol in $DLPPolicy) {
        "<h3>$($DLPPol.DisplayName)</h3><div class='h3-content'>" | Add-Content $HTML
        if ([string]::IsNullOrEmpty($DLPPol.CreatedBy)) { $DLPPol.CreatedBy = "Microsoft Corporation" }        
        "<div style='background-color: #f8f9fa; padding: 10px; border-radius: 4px; margin: 0 0 8px 0; color: #212529;'><Strong>Created by:</Strong> $($DLPPol.CreatedBy)<br><Strong>Last modified by: </Strong>$($DLPPol.LastModifiedBy)<br><Strong>Last modified:</Strong> $($DLPPol.WhenChanged.ToString('yyyy-MM-dd'))</div>" | Add-Content $HTML
        if ([string]::IsNullOrEmpty($DLPPol.Comment)) { '<div style="background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;"><strong>Description:</strong> No description provided</div>' | Add-content $HTML }
        else { "<div style='background-color: #e3f2fd; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Description: </strong>$($DLPPol.Comment)</div>" | Add-content $HTML }
        "<div style='background-color: #e8f5e8; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>Mode:</strong> $($polMode[$DLPPol.Mode])</div>" | Add-content $HTML


        #AdminUnits
        #ADMINUNITS



        #Locations
        #Location table
        "<table><tr><th>Location</th><th>Assigned to:</th></tr>" | Add-Content $HTML

        #Exchange
        $DLPExcInc = $DLPPol.ExchangeLocation
        $DLPExcExConv = $DLPPol.ExchangeSenderMemberOfException | ConvertFrom-Json 
        $DLPExcMbOf = $DLPPol.ExchangeSenderMemberOf | ConvertFrom-Json 
        if ($DLPExcInc.DisplayName) {
            if ($DLPExcMbOf.DisplayName) {
                "<tr><td>Exchange Location</td><td>" | Add-Content $HTML
                foreach ($group in $DLPExcMbOf) { "$($group.DisplayName) <strong>($($RecipientTypeConv[[string]$group.RecipientTypeDetails]))</strong><br>" | Add-Content $HTML }
                "</td></tr>" | Add-Content $HTML
            }
            else { "<tr><td>Exchange Location</td><td>$($DLPExcInc.DisplayName)</td></tr>" | Add-Content $HTML }
            if ($DLPExcExConv.DisplayName) {
                "<tr><td>Exchange Exception</td><td>" | Add-Content $HTML
                foreach ($group in $DLPExcExConv) { "$($group.DisplayName) <strong>($($RecipientTypeConv[[string]$group.RecipientTypeDetails]))</strong><br>" | Add-Content $HTML }
                "</td></tr>" | Add-Content $HTML
            }
            else { "<tr><td>Exchange Exclusion</td><td>N/A</td></tr>" | Add-Content $HTML }
        }

        #SharePoint Location
        $DLPSPOInc = $DLPPol.SharePointLocation
        $DLPSPOExConv = $DLPPol.SharePointLocationException 
        $DLPSPOASInc = $DLPPol.SharePointAdaptiveScopes 
        $DLPSPOASExc = $DLPPol.SharePointAdaptiveScopesException
        if ($DLPSPOInc.DisplayName -or $DLPSPOASInc.DisplayName) {
            "<tr><td>SharePoint Location</td><td>" | Add-Content $HTML
            if ($DLPSPOInc.DisplayName -eq "All") { "All" | Add-Content $HTML }
            elseif ($DLPSPOInc.DisplayName) {
                "<strong>SPO Sites:</strong><br>" | Add-Content $HTML
                "$($DLPSPOInc.DisplayName -join "<br>") <br>" | Add-Content $HTML 
            }
            if ($DLPSPOASInc.DisplayName) {
                "<strong>Adaptive Scopes:</strong><br>" | Add-Content $HTML
                ($DLPSPOASInc.DisplayName -join "<br>") | Add-Content $HTML
            }
            "</td></tr>" | Add-Content $HTML
            if ($DLPSPOExConv -or $DLPSPOASExc) {
                "<tr><td>SharePoint Exclusion</td><td>" | Add-Content $HTML
                if ($DLPSPOExConv.DisplayName) {
                    "<strong>SPO Sites:</strong><br>" | Add-Content $HTML
                    "$($DLPSPOExConv.DisplayName -join "<br>") <br>" | Add-Content $HTML
                }
                if ($DLPSPOASExc.DisplayName) {
                    "<strong>Adaptive Scopes:</strong><br>" | Add-Content $HTML
                    ($DLPSPOASExc.DisplayName -join "<br>") | Add-Content $HTML
                }
                "</td></tr>" | Add-Content $HTML
            }
            else { "<tr><td>SharePoint Exclusion</td><td>N/A</td></tr>" | Add-Content $HTML }
        }

        #OneDrive Location
        $DLPODLoc = $DLPPol.OneDriveLocation
        $DLPODInc = $DLPPol.OneDriveSharedBy + $DLPPol.OneDriveSharedByMemberOf | ConvertFrom-Json
        $DLPSPOASExc = $DLPPol.ExceptIfOneDriveSharedByMemberOf + $DLPPol.ExceptIfOneDriveSharedBy | ConvertFrom-Json
        if ($DLPODLoc) {
            if ($DLPODInc.DisplayName) {
                "<tr><td>OneDrive Location</td><td>" | Add-Content $HTML
                foreach ($group in $DLPODInc) { "$($group.DisplayName) <strong>($($RecipientTypeConv[[string]$group.RecipientTypeDetails]))</strong><br>" | Add-Content $HTML }
                "</td></tr>" | Add-Content $HTML 
            }
            else { "<tr><td>OneDrive Location</td><td>$($DLPODLoc.DisplayName)</td></tr>" | Add-Content $HTML }
            if ($DLPSPOASExc.DisplayName) {
                "<tr><td>OneDrive Exception</td><td>" | Add-Content $HTML
                foreach ($group in $DLPSPOASExc) { "$($group.DisplayName) <strong>($($RecipientTypeConv[[string]$group.RecipientTypeDetails]))</strong><br>" | Add-Content $HTML }
                "</td></tr>" | Add-Content $HTML
            }
            else { "<tr><td>OneDrive Exclusion</td><td>N/A</td></tr>" | Add-Content $HTML }
    
        }
        #Teams Location

        $teamsGroups = @{
            "IndividualResource" = "User"
            "Group"              = "Group"
        }

        $DLPTmsLoc = $DLPPol.TeamsLocation
        $DLPTmsExc = $DLPPol.TeamsLocationException
        if ($DLPPol.TeamsLocation) {
            if ($DLPTmsLoc.DisplayName -eq "All") { "<tr><td>Teams Location</td><td>$($DLPTmsLoc.DisplayName)</td></tr>" | Add-Content $HTML }
            else {
                "<tr><td>Teams Location</td><td>" | Add-Content $HTML  
                foreach ($assign in $DLPTmsLoc) { "$($assign.DisplayName) <strong>($($teamsGroups[[string]$assign.Type]))</strong><br>" | Add-Content $HTML }
                "</td></tr>" | Add-Content $HTML 
            }
            if ($DLPTmsExc.DisplayName) {
                "<tr><td>Teams Exception</td><td>" | Add-Content $HTML
                foreach ($assign in $DLPTmsExc) { "$($assign.DisplayName) <strong>($($teamsGroups[[string]$assign.Type]))</strong><br>" | Add-Content $HTML }
                "</td></tr>" | Add-Content $HTML 
            }
            else { "<tr><td>Teams Exclusion</td><td>N/A</td></tr>" | Add-Content $HTML }
        }
        #Devices - user and groups
        $DLPDevLoc = $DLPPol.EndpointDlpLocation
        $DLPDevExc = $DLPPol.EndpointDlpLocationException
        $DLPDevAs = $DLPPol.EndpointDlpExtendedLocations | ConvertFrom-Json
        $DLPDevInc = $DLPDevAs.Inclusions
        $DLPDevExcl = $DLPDevAs.Exclusions
        if ($DLPDevLoc.DisplayName) {
            "<tr><td>Device Location</td><td>" | Add-Content $HTML
            if ($DLPDevLoc.DisplayName -eq "All" -and $DLPDevLoc) { "<strong>User/User Group:</strong><br>All<br>" | Add-Content $HTML }
            elseif ($DLPDevLoc.DisplayName -eq "All") { "<strong>User/User Group:</strong>All" | Add-Content $HTML }
            else {
                "<strong>User/User Group:</strong><br>" | Add-Content $HTML
                foreach ($group in $DLPDevLoc) {
                    "$($group.DisplayName) $(if ($group.Type.ToString() -eq 'Group') { '(Group)' })<br>" | Add-Content $HTML
                }
            }
            if ($DLPDevInc) {
                "<strong>Device/Device Group:</strong><br>" | Add-Content $HTML
                foreach ($group in $DLPDevInc) { "$($group.DisplayName) $(if ($group.Type.ToString() -eq 'Group') { '(Group)' })<br>" | Add-Content $HTML }
            }
            "</td></tr>" | Add-Content $HTML
            "<tr><td>Device Exclusion</td><td>" | Add-Content $HTML
   
            if ($DLPDevExc.DisplayName -or $DLPDevExcl) { 
                "<strong>User/User Group:</strong><br>" | Add-Content $HTML
                if ($DLPDevExc) { foreach ($group in $DLPDevExc) { "$($group.DisplayName) $(if ($group.Type.ToString() -eq 'Group') { '(Group)' })<br>" | Add-Content $HTML } }
                if ($DLPDevExcl) {
                    "<strong>Device/Device Group:</strong><br>" | Add-Content $HTML
                    foreach ($group in $DLPDevExcl) { "$($group.DisplayName) $(if ($group.Type.ToString() -eq 'Group') { '(Group)' })<br>" | Add-Content $HTML }
                }
                "</td></tr>" | Add-Content $HTML
            }
            else { "N/A</tr>" | Add-Content $HTML }
        }

        #Instances - WILL NEED LOOKING AT, AT A LATER DATE WHEN EXAMPLES CAN BE OBSERVED ####################################################################################################################
        if ($DLPPol.ThirdPartyAppDlpLocation.DisplayName -eq "All") {
            "<tr><td>Instances</td><td>All Instaces</td></tr>" | Add-Content $HTML
        }
        elseif (-not([string]::IsNullOrEmpty($DLPPol.ThirdPartyAppDlpLocation.DisplayName)) -and $DLPPol.ThirdPartyAppDlpLocation.DisplayName -notcontains "All") {
            "<tr><td>Instances</td><td>Specific Instances have been selected</td></tr>" | Add-Content $HTML
        }
        #Instances - WILL NEED LOOKING AT, AT A LATER DATE WHEN EXAMPLES CAN BE OBSERVED ####################################################################################################################
        
        #On-Prem
        if ($DLPPol.OnPremisesScannerDlpLocation) {
            "<tr><td>On-premises repositories</td><td>$(($DLPPol.OnPremisesScannerDlpLocation.DisplayName) -join '<br>')</td></tr>" | Add-Content $HTML
            if ($DLPPol.OnPremisesScannerDlpLocationException) { "<tr><td>On-premises repositories exclusions</td><td>$(($DLPPol.OnPremisesScannerDlpLocationException.DisplayName) -join '<br>')</td></tr>" | Add-Content $HTML }
            else { "<tr><td>On-premises repositories exclusions</td><td>N/A</td></tr>" | Add-Content $HTML }
        }

        #End of location table
        "</table>" | Add-Content $HTML

        ################################################### DLP RULES ################################################


        foreach ($DLPPoll in $DLPPol.Name) {
            $DLPRules = Get-DlpComplianceRule -Policy $DLPPoll
            foreach ($DLPRule in $DLPRules.Name) {
                $DLPRuleInfo = Get-DlpComplianceRule -Identity $DLPRule
                "<div style='background-color: #e8f5e8; padding: 10px; border-radius: 4px; margin: 8px 0; color: #212529;'><strong>$($DLPRuleInfo.DisplayName):</strong></div>" | Add-Content -Path $HTML                
                $DLPAdvRule = $DLPRuleInfo.AdvancedRule | ConvertFrom-Json
                $DLPAdvRuleVal = $DlPAdvRule.Condition.SubConditions
                foreach ($condition in $DLPAdvRuleVal) {
                    if ($condition.ConditionName -eq "ContentContainsSensitiveInformation") {
                        "<strong>$($condition.ConditionName)</strong><br>" | Add-Content $HTML
                        foreach ($sensitiveType in $condition.Value.Groups.Sensitivetypes) {
                            "$($sensitiveType.Name) --- $($sensitiveType.ConfidenceLevel) --- $($sensitiveType.MinCount) to $($sensitiveType.MaxCount)<br>" | Add-Content $HTML
                        }
                    }
                    else {
                        "<Strong>$($condition.ConditionName)</Strong><br>$($condition.Value)<br>" | Add-Content $HTML
                    }   
                }
            }
        }
        #Close section-content div
        "</div>" | Add-Content $HTML
    }
    #Close section-content div
    "</div>" | Add-Content $HTML


}


##########################################################################################################################################################


function retentionPolicies {

    Write-Host "Gathering Retention Policies"
    #RetentionPolicies
    '<h2>Retention Policies</h2><div class="section-content">' | Add-Content -Path $HTML

    ##########################################################################################################################################################

    $RPA = Get-RetentionCompliancePolicy -DistributionDetail -RetentionRuleTypes | Where-Object { $_.RetentionRuleTypes -contains "Apply" -or $_.RetentionRuleTypes -contains "Default" }

    $RPD = $RPA | Where-Object { $_.RetentionRuleTypes -contains "Apply" -or $_.RetentionRuleTypes -contains "Default" }
    Write-Host "Found $($RPD.Count) retention policies"

    foreach ($RP in $RPD) {
        "<h3>$($RP.Name)</h3><div class='h3-content'>" | Add-Content -Path $HTML

        #Close section-content div
        "</div>" | Add-Content $HTML
    }
    #Close section-content div
    "</div>" | Add-Content $HTML

}


##########################################################################################################################################################

function retentionLabels {

    Write-Host "Gathering Retention Labels"
    #Labels
    '<h2>Retention Labels</h2><div class="section-content">' | Add-Content -Path $HTML

    ##########################################################################################################################################################

    $RLD = $RPA | Where-Object { $_.RetentionRuleTypes -eq "Publish" }
    Write-Host "Found $($RLD.Count) retention labels"

    foreach ($RL in $RLD) {
        "<h3>$($RP.Name)</h3><div class='h3-content'>" | Add-Content -Path $HTML

        #Close section-content div
        "</div>" | Add-Content $HTML
    }
    #Close section-content div
    "</div>" | Add-Content $HTML

}


##########################################################################################################################################################

function tenantControls {

    Write-Host "Gathering Tenant Controls"
    #Labels
    '<h2>Tenant Controls</h2><div class="section-content">' | Add-Content -Path $HTML

    ##########################################################################################################################################################

    #Co-Authoring
    Write-Host "Checking Co-Authoring setting..."
    $CoAuthoring = (Get-PolicyConfig).EnableLabelCoauth
    if ($CoAuthoring -eq "True") {
        "<p>Co-Authoring is enabled</p>" | Add-Content $HTML
    }
    else {
        "<p>Co-Authoring is disabled</p>" | Add-Content $HTML
    }

    #Exchange roles
    Write-Host "Checking Exchange roles..."
    $ExchangeRoles = Get-ManagementRoleAssignment -RoleAssigneeType "RoleAssignmentPolicy" |
    Where-Object {
        $_.Role -eq "My Custom Apps" -or
        $_.Role -eq "My Marketplace Apps" -or
        $_.Role -eq "My ReadWriteMailbox Apps"
    }

    if ($ExchangeRoles) {
        $ExchangeRoles |
        Group-Object RoleAssigneeName |
        ForEach-Object {
            $name = $_.Name
            $roles = $_.Group.Role -join ", "
            "<p>$name contains: $roles</p>" | Add-Content $HTML
        }
    }
    else {
        "<p>There are no Exchange user roles with 'My Custom Apps', 'My Marketplace Apps', or 'My ReadWriteMaibox Apps'</p>" | Add-Content $HTML
    }

    #CustomerLockbox
    Write-Host "Checking Customer Lockbox..."
    $OrganizationConfig = Get-OrganizationConfig | export-csv -Path "C:\temp\test2.csv"

    if ($OrganizationConfig.CustomerLockBoxEnabled -eq $False) {
        "<p>Customer Lockbox is disabled</p>" | Add-Content $HTML
    }
    else {
        "<p>Customer Lockbox is enabled</p>" | Add-Content $HTML
    }


    #OfficeApps
    Write-Host "Checking Office Apps settings..."
    $Uri = "https://graph.microsoft.com/beta/admin/appsAndServices/settings"
    Invoke-MgGraphRequest -Uri $Uri

    #Close section-content div
    "</div>" | Add-Content -Path $HTML

    #Teams cloud apps
    #Not currently possible in the same script

}


##########################################################################################################################################################

#Get-RetentionComplianceRule -Policy
#Get-LabelPolicyRule | select LabelName

##########################################################################################################################################################

'<div id="InformationProtection" class="tab-content active">' | Add-Content $HTML

labels
labelPolicies
AutoLabelingPolicy

'</div><div id="DLP" class="tab-content">' | Add-Content $HTML

DLPPolicies

#Finalise and start

@"
</div>
<script>

function openTab(evt, tabName) {

    document.querySelectorAll(".tab-content").forEach(tab => {
        tab.classList.remove("active");
    });

    document.querySelectorAll(".tab-button").forEach(btn => {
        btn.classList.remove("active");
    });

    document.getElementById(tabName).classList.add("active");
    evt.currentTarget.classList.add("active");
}

function collapseAll() {
    // Collapse all h2 sections
    document.querySelectorAll("h2").forEach(h2 => {
        h2.classList.remove("active");
        const content = h2.nextElementSibling;
        if (content && content.classList.contains('section-content')) {
            content.style.display = "none";
        }
    });

    // Collapse all h3 sections
    document.querySelectorAll("h3").forEach(h3 => {
        h3.classList.remove("active");
        const content = h3.nextElementSibling;
        if (content && content.classList.contains('h3-content')) {
            content.style.display = "none";
        }
    });

    // Collapse all details elements (label groups)
    document.querySelectorAll("details").forEach(details => {
        details.removeAttribute("open");
    });

    // Update the arrow symbols for h2 and h3
    document.querySelectorAll("h2::after").forEach(arrow => {
        arrow.content = " ▼";
    });
    document.querySelectorAll("h3::after").forEach(arrow => {
        arrow.content = " ▼";
    });
}

function expandAll() {
    // Expand all h2 sections
    document.querySelectorAll("h2").forEach(h2 => {
        h2.classList.add("active");
        const content = h2.nextElementSibling;
        if (content && content.classList.contains('section-content')) {
            content.style.display = "block";
        }
    });

    // Expand all h3 sections
    document.querySelectorAll("h3").forEach(h3 => {
        h3.classList.add("active");
        const content = h3.nextElementSibling;
        if (content && content.classList.contains('h3-content')) {
            content.style.display = "block";
        }
    });

    // Expand all details elements (label groups)
    document.querySelectorAll("details").forEach(details => {
        details.setAttribute("open", "");
    });

    // Update the arrow symbols for h2 and h3
    document.querySelectorAll("h2::after").forEach(arrow => {
        arrow.content = " ▲";
    });
    document.querySelectorAll("h3::after").forEach(arrow => {
        arrow.content = " ▲";
    });
}

document.querySelectorAll("h2").forEach(h2 => {
    h2.addEventListener("click", () => {
        h2.classList.toggle("active");
        const content = h2.nextElementSibling;
        content.style.display =
            (content.style.display === "block") ? "none" : "block";
    });
});

document.querySelectorAll("h3").forEach(h3 => {
    h3.addEventListener("click", () => {
        h3.classList.toggle("active");
        const content = h3.nextElementSibling;
        content.style.display =
            (content.style.display === "block") ? "none" : "block";
    });
});

</script>

</body>
</html>
"@ | Add-Content -Path $HTML

Start-Process -FilePath $HTML
