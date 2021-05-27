# Summary
This script takes an SCCM collection name, the name of an application (product), and the filepath of a text file. It polls each computer in the collection and returns a table containing the version of the product installed on each computer. Everything output to the screen is also sent to the log file.  

Must be run on a computer with the SCCM admin console installed if using `-Collection` parameter.  

# Usage
1. Download `Report-ProductVersion.psm1` to `$HOME\Documents\WindowsPowerShell\Modules\Report-ProductVersion\Report-ProductVersion.psm1`.
2. Open a PowerShell console as the user which has SCCM permissions (For Engineering, this is probably your regular NetID and NOT your SU account).
    - If this account differs from your regular account, you may need to explicitly import the module: `Import-Module "c:\path\to\Report-ProductVersion.psm1`
3. Run the command:
  - e.g. `Report-ProductVersion -Collection "UIUC-ENGR-ESPL" -Product "*acrobat*" -Log "c:\espl-acrobat.log" -Csv "c:\espl-acrobat.csv"`
  - e.g. `Report-ProductVersion -Computers "gelib-4e-01","eh-406b1-*","mel-1001-01" -Product "*acrobat*"`

# Example
```powershell
C:\> Report-ProductVersion -Collection "UIUC-ENGR-ESPL" -Product "*acrobat*" -Log "c:\epsl-acrobat.log" -Csv "c:\espl-acrobat.csv"

Computers in collection "UIUC-ENGR-ESPL":
ESPL-114-01
ESPL-114-02
ESPL-114-03
ESPL-114-04
ESPL-114-05
ESPL-114-06
ESPL-114-07
ESPL-114-08
ESPL-114-09
ESPL-MACH-02

Polling computer "ESPL-114-01"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20034 20.006.20042".
Polling computer "ESPL-114-02"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20034 20.009.20063".
Polling computer "ESPL-114-03"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20034 20.006.20042".
Polling computer "ESPL-114-04"...
    Success. Name(s): "Acrobat DC 19.012.20036 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20036 20.009.20063".
Polling computer "ESPL-114-05"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20034 20.006.20042".
Polling computer "ESPL-114-06"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20034 20.009.20063".
Polling computer "ESPL-114-07"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20034 20.006.20042".
Polling computer "ESPL-114-08"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC", Version(s): "1.0.0000 19.012.20034 20.009.20063".
Polling computer "ESPL-114-09"...
    Unknown COMException!
Polling computer "ESPL-MACH-02"...
    Success. Name(s): "Acrobat DC 19.012.20034 SDL ITP Adobe Acrobat DC Adobe Acrobat Reader DC Acrobat", Version(s): "1.0.0000 19.012.20034 20.006.20042 1.0.0000".

PSComputerName Name                            Version
-------------- ----                            -------
ESPL-114-01    Adobe Acrobat Reader DC         20.006.20042
ESPL-114-01    Adobe Acrobat DC                19.012.20034
ESPL-114-01    Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-114-02    Adobe Acrobat Reader DC         20.009.20063
ESPL-114-02    Adobe Acrobat DC                19.012.20034
ESPL-114-02    Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-114-03    Adobe Acrobat Reader DC         20.006.20042
ESPL-114-03    Adobe Acrobat DC                19.012.20034
ESPL-114-03    Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-114-04    Adobe Acrobat Reader DC         20.009.20063
ESPL-114-04    Adobe Acrobat DC                19.012.20036
ESPL-114-04    Acrobat DC 19.012.20036 SDL ITP 1.0.0000
ESPL-114-05    Adobe Acrobat Reader DC         20.006.20042
ESPL-114-05    Adobe Acrobat DC                19.012.20034
ESPL-114-05    Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-114-06    Adobe Acrobat Reader DC         20.009.20063
ESPL-114-06    Adobe Acrobat DC                19.012.20034
ESPL-114-06    Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-114-07    Adobe Acrobat Reader DC         20.006.20042
ESPL-114-07    Adobe Acrobat DC                19.012.20034
ESPL-114-07    Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-114-08    Adobe Acrobat Reader DC         20.009.20063
ESPL-114-08    Adobe Acrobat DC                19.012.20034
ESPL-114-08    Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-114-09    Error                           Error
ESPL-MACH-02   Adobe Acrobat Reader DC         20.006.20042
ESPL-MACH-02   Acrobat                         1.0.0000
ESPL-MACH-02   Acrobat DC 19.012.20034 SDL ITP 1.0.0000
ESPL-MACH-02   Adobe Acrobat DC                19.012.20034

EOF

C:\>
```

# Parameters

### -Computers [string array]
Required string array, if not using `-Collection`.  
An array of strings representing either exact computer names, computername queries using `*` wildcards, or a combination thereof.  
e.g. `Report-ProductVersion -Computers "gelib-4e-01","eh-406b1-*","mel-1001-01" -Product "*acrobat*"`

### -Collection [string]
Required string.  
The MECM collection containing the computers you want to query.  

### -Product [string]
Required string.  
The name query of the product you want to query for.  
The script will search for all products whose `Name` or `DisplayName`(depending on the WMI data source) exactly match the given product string. Use `*` as a wildcard.  

### -Log [string]
Optional string.  
The full path to a log file where progress will be logged.  
If omitted, no log will be created.  

### -Csv [string]
Optional string.  
The full path to a CSV file where the results will be logged.  
If omitted, no CSV will be created.  

### -LogIncrementalProgress
Optional switch.  
If specified, a table of cumulative result data will be logged each time a computer is polled. Useful in case the script is getting hung up, so you get at least some data.  

### -SiteCode [string]
Optional string.  
The SiteCode of the MECM site to use.  
Default is `MP0`.  

### -Provider [string]
Optional string.  
The ProviderMachineName of the MECM server to use.  
Default is `sccmcas.ad.uillinois.edu`.  

# Notes
- For large collections, this can take a long time. When machines cannot be contacted it takes 20-25 seconds for the query to timeout, occasionally hanging for much longer. Successful queries may still take a while.
- To be the most thorough, the script queries WMI for both `Win32_Product` and `Win32Reg_AddRemovePrograms`, as these can both provide distinct sets of results.  
- By mseng3. See my other projects here: https://github.com/mmseng/code-compendium.
