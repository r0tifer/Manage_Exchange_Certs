@{
    # Module Metadata
    RootModule           = 'ManageExchangeCert.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = 'b8f2d7c4-75ff-4bc9-8b30-2d7a9e945f3f'
    Author               = 'Michael Levesque'
    CompanyName          = ''
    Copyright            = 'GNU'
    Description          = 'Manage Exchange Certificates: Renew, Replace, Import, Export, Revoke, and Auto-Assign with CLI support.'

    # Module Components
    FunctionsToExport    = @(
        'Start-CertManager',
        'Renew-Certificate',
        'Replace-Certificate',
        'Revoke-Certificate',
        'Delete-Certificate',
        'Export-Certificate',
        'Import-Certificate',
        'Resume-PendingCertificate',
        'Select-ExchangeServer',
        'Prompt-CertSelection',
        'Show-MainMenu',
        'Show-CertList'
    )

    CmdletsToExport      = @()
    VariablesToExport    = '*'
    AliasesToExport      = @()

    # Compatibility
    PowerShellVersion    = '5.1'

    # Nested modules (if any)
    NestedModules        = @()

    # Optional script files (none here)
    ScriptsToProcess     = @()

    # Private data (extend if you want to support CLI argument parsing later)
    PrivateData = @{
        PSData = @{
            Tags       = @('Exchange', 'Certificate', 'CLI', 'Renewal', 'PowerShell')
            ProjectUri = 'https://github.com/r0tifer/Manage_Exchange_Certs'
        }
    }
}
