@{
    ExcludeRules = @(
        # Parse-Receipt uses a non-standard verb intentionally -- it is an internal
        # helper, not an exported cmdlet, and the name is the clearest description.
        'PSUseApprovedVerbs',

        # Get-ValidAccounts, Get-Categories, Set-CategoryNamedRanges intentionally
        # use plural nouns because they operate on collections, not single items.
        'PSUseSingularNouns',

        # This is a console utility, not a module. Write-Host output is the intended
        # user-facing log. Write-Output would pollute the pipeline incorrectly.
        'PSAvoidUsingWriteHost',

        # Empty catch blocks are used intentionally as silent COM probes, e.g.
        # try { $sheet = $wb.Sheets.Item("Name") } catch {}
        # where a missing sheet is handled by checking for $null on the next line.
        'PSAvoidUsingEmptyCatchBlock',

        # Set-CategoryNamedRanges and Set-SubcategoryValidationXml are internal helpers
        # that change state by design. ShouldProcess (-WhatIf/-Confirm) is not
        # appropriate for a non-interactive batch script.
        'PSUseShouldProcessForStateChangingFunctions'
    )
}
