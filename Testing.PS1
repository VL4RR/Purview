$allConditionsCondition = @{
    "ContentContainsSensitiveInformation"  = "Content Contains:"
    "AccessScope"                          = "Content is shared from Microsoft 365"
    "SenderIPRanges"                       = "Sender IP address is:"
    "RecipientDomainIs"                    = "Recipient domain is:"
    "ContentExtensionMatchesWords"         = "File extension is:"
    "DocumentIsPasswordProtected"          = "Document or attachment is password protected"
    "DocumentIsUnsupported"                = "Document could not be scanned"
    "ProcessingLimitExceeded"              = "Document didn't complete scanning"
    "HeaderContainsWords"                  = "Header contains words of phrases:"
    "SubjectMatchesPatterns"               = "Subject matches patterns:"
    "AnyOfRecipientAddressContainsWords"   = "Recipient address contains words:"
    "FromAddressMatchesPatterns"           = "Sender address matches patterns:"
    "SenderDomainIs"                       = "Sender domain is:"
    "SentToMemberOf"                       = "Recipient is a member of:"
    "From"                                 = "Sender is:"
    "DocumentNameMatchesWords"             = "Document name contains words or phrases:"
    "ContentPropertyContainsWords"         = "Document property is:"
    "DocumentSizeOver"                     = "Document size equals or is greater than:"
    "MessageSizeOver"                      = "Message size equals or is greater than:"
    "SentTo"                               = "Recipient is:"
    "ContentIsNotLabeled"                  = "Content is not labeled"
    "FromScope"                            = "Content is recieved from:"
    "ContentCharacterSetContainsWords"     = "Content character set contains words:"
    "HeaderMatchesPatterns"                = "Header matches patterns:"
    "SenderADAttributeMatchesPatterns"     = "Sender AD Attribute matches patterns:"
    "RecipientADAttributeContainsWords"    = "Recipient AD Attribute contains words or phrases:"
    "RecipientADAttributeMatchesPatterns"  = "Recipient AD Attribute matches patterns:"
    "HasSenderOverride"                    = "Has sender overridden the policy tip"
    "FromMemberOf"                         = "Sender is a member of:"
    "DocumentScanFailures"                 = "Document scan failures:"
    "AnyOfRecipientAddressMatchesPatterns" = "Recipient address matches patterns:"
    "DocumentNameMatchesPatterns"          = "Document name matches patterns:"
    "SubjectContainsWords"                 = "Subject contains words or phrases:"
    "SubjectOrBodyContainsWords"           = "Subject or body contains words or phrases:"
    "SubjectOrBodyMatchesPatterns"         = "Subject or body matches patterns:"
    "FromAddressContainsWords"             = "Sender address contains words:"
    "MessageTypeMatches"                   = "Message type is:"
    "WithImportance"                       = "Message Importance is:"
    "RecipientCountOver"                   = "Unique recipients greater than:"
    "DomainCountOver"                      = "Unique domains greater than:"
    "SenderADAttributeContainsWords"       = "Sender AD Attribute contains words or phrases:"
    "DocumentContainsWords"                = "Document content contains words or phrases:"
    "DocumentMatchesPatterns"              = "Document content matches patterns:"

    "DocumentCreatedBy"                    = "Document created by:"
    "DocumentCreatedOnOrAfter"             = "Document creation date is on or after"
    "DocumentCreatedOnOrBefore"            = "Document creation date is on or before"
    "DocumentLastModifiedOnOrAfter"        = "Document modification date is on or after"
    "DocumentLastModifiedOnOrBefore"       = "Document modification date is on or before"
    "ContentIsSharedWith"                  = "Document is shared:"
}

$Labels = Get-Label
$script:Labels = $Labels

function SensitiveTypes {
    param($Condition)

    $sensitiveTypeTC = $Condition.Value.Groups.Sensitivetypes | Where-Object { $_.Classifiertype -eq "MLModel" -and $_.Name }
    $sensitiveTypeSIT = $Condition.Value.Groups.Sensitivetypes | Where-Object { $_.Classifiertype -ne "MLModel" -and $_.Name }
    $sensitiveTypeLabel = $Condition.Value.Groups.Labels | Where-Object { $_.Name }
    Write-Host "$($allConditionsCondition[$Condition.ConditionName])"
    if ($sensitiveTypeTC) {
        Write-Host "Trainable Classifiers:"
        foreach ($sensitiveType in $sensitiveTypeTC) { Write-Host "$($sensitiveType.Name)" }
        Write-Host ""
    }
    if ($sensitiveTypeSIT) {
        Write-Host "Sensitive Information Types:"
        foreach ($sensitiveType in $sensitiveTypeSIT) {
            Write-Host "$($sensitiveType.Name) --- $($sensitiveType.ConfidenceLevel) --- $($sensitiveType.MinCount) to $($sensitiveType.MaxCount)"
        }
        Write-Host ""
    }
    if ($sensitiveTypeLabel) {
        Write-Host "Labels:"
        foreach ($sensitiveType in $sensitiveTypeLabel) {
            $label = $script:Labels | Where-Object { $_.Guid -eq $sensitiveType.Id } | Select-Object -First 1
            if ($label) { Write-Host $label.DisplayName }
            else { Write-Host $sensitiveType.Name }
        }
        Write-Host ""
    }
}

function PSObjectConditions {
    param($Condition)
    if ($($allConditionsCondition[$Condition.ConditionName])) { Write-Host "$($allConditionsCondition[$Condition.ConditionName])" } else { Write-Host "$($Condition.ConditionName):" }
    Write-host "$($Condition.Value.PSObject.Properties.Name)`n$($Condition.Value.PSObject.Properties.Value -join "`n")`n"
}


$DLPPol = Get-DlpCompliancePolicy -Identity "ComplexTest"

foreach ($DLPPoll in $DLPPol.Name) {
    $DLPRules = Get-DlpComplianceRule -Policy $DLPPoll
    foreach ($DLPRule in $DLPRules.Name) {
        $DLPRuleInfo = Get-DlpComplianceRule -Identity $DLPRule
        Write-Host "$($DLPRuleInfo.DisplayName)`n" -ForegroundColor Green
        $DLPAdvRule = $DLPRuleInfo.AdvancedRule | ConvertFrom-Json
        $DLPAdvRuleVal = $DLPAdvRule.Condition.SubConditions
        if ($DLPAdvRuleVal.Operator.Count -ge 1) { 
            Write-Host "Group 1:`n"  
            $groupNumber = 1 
        }
        foreach ($condition in $DLPAdvRuleVal) {
            if ($condition.ConditionName -eq "ContentContainsSensitiveInformation") { SensitiveTypes $condition }
            elseif ($condition.Value -is [System.Management.Automation.PSObject] -and $GroupCondition.ConditionName -ne "ContentContainsSensitiveInformation") { PSObjectConditions $condition }
            elseif ($condition.Operator -eq "And") {
                $groupNumber ++
                Write-Host "Group $groupNumber`n"
                foreach ($GroupCondition in $condition.SubConditions) {
                    if ($GroupCondition.ConditionName -eq "ContentContainsSensitiveInformation") { SensitiveTypes $GroupCondition }
                    elseif ($GroupCondition.Value -is [System.Management.Automation.PSObject] -and $GroupCondition.ConditionName -ne "ContentContainsSensitiveInformation") { PSObjectConditions $GroupCondition }
                    else { Write-Host "$($GroupCondition.ConditionName)`n$($GroupCondition.Value)`n" }
                }
            }
            else { if ($($allConditionsCondition[$condition.ConditionName])) { Write-Host "$($allConditionsCondition[$condition.ConditionName])`n$($condition.Value -join "`n")`n" } else { Write-Host "$($condition.ConditionName)`n$($condition.Value -join "`n")`n" } }
        }
    }
}

#$allConditionsCondition[$DLPAdvRuleVal.ConditionName]
#$DLPAdvRuleVal.ConditionName.Count

