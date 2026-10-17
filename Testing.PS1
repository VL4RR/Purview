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

function SensitiveTypes {
    param($Condition)

    Write-Host "$($Condition.ConditionName)"

    foreach ($sensitiveType in $Condition.Value.Groups.Sensitivetypes) {
        Write-Host "$($sensitiveType.Name) --- $($sensitiveType.ConfidenceLevel) --- $($sensitiveType.MinCount) to $($sensitiveType.MaxCount)"
    }
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
            Write-Host "Group 1:"  
            $groupNumber = 1 
        }
        foreach ($condition in $DLPAdvRuleVal) {
            if ($condition.ConditionName -eq "ContentContainsSensitiveInformation") { SensitiveTypes $condition }
            elseif ($condition.Operator -eq "And") {
                $groupNumber ++
                Write-Host "Group $groupNumber"
                foreach ($GroupCondition in $condition.SubConditions) {

                    if ($GroupCondition.ConditionName -eq "ContentContainsSensitiveInformation") { SensitiveTypes $GroupCondition }
                    else { Write-Host "$($GroupCondition.ConditionName)`n$($GroupCondition.Value)`n" }
                }
            }
            else { Write-Host "$($condition.ConditionName)`n$($condition.Value)`n" }
        }
    }
}

#$DLPAdvRule.Condition.SubConditions.Subconditions | Format-List *
$DLPAdvRuleVal.Value.Groups.Sensitivetypes | format-List *
## Need to sort labels