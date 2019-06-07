# CV-PstCopy
PowerShell to get PST files copied from one folder to another in the correct folder format for Commvault via extraction of username from filename. It is specific to Transvault output which is in the format Recipients.UserName`@`domain.com_Number.pst

* Scripts purpose is to copy or move PST's from source folder to target folder structure necessary for CommVault PST import

* If PST file owner set correctly then script is not necessary.

* Creates log file

    - Run As Administrator! 
    - Scripts purpose is to move PST’s from source folder to target folder structure necessary for Commvault PST archive user association”
    - If PST file owner set correctly then script is not necessary
    - Currently specific to only copying PSTs from TransVault migration output from Enterprise Vault.
    - Creates log file if doesn’t exist in same directory as script
    - Creates hashsum for file prior and after copying and compares to determine if mismatch
    - Renames source files after copy to indicate they were copied. Not applicable with -move option
    - source and target directory must exist
    - -source = Source directory where PST’s are located.
    - -target = Target main directory where PST’s are to be moved to
    - -move is an optional parameter that moves rather than copies PST's so faster and saves space.

    Example Usage:
    
    To copy PST files:
    
        .\CV-PstCopy.ps1 -source C:\scripts\test\srcdir -target C:\scripts\test\trgtdir
        
    To move PST files:
        
        .\CV-PstCopy.ps1 -source C:\scripts\test\srcdir -target C:\scripts\test\trgtdir -move
