@{
    # Log configuration
    LogDir = '.\Logs'
    MaxLogSizeMB = 10
    EnableTranscript = $true
    
    # AD replication settings
    ReplicationTimeoutSeconds = 60
    ReplicationRetryIntervalSeconds = 5
    
    # Event processing settings
    EventLookupMinutes = 5
    MaxEvents = 500
    
    # Computer routing rules
    # Pattern: Regular expression to match computer names
    # OU: Target Organizational Unit distinguished name
    # Description: Human-readable description (optional)
    Rules = @(
        @{ 
            Pattern = '^Alpha'
            OU = 'OU=Location Alpha,OU=Computers,OU=DEV Company,DC=dev,DC=local'
            Description = 'Alpha location computers'
        }
        @{ 
            Pattern = '^Beta'
            OU = 'OU=Location Beta,OU=Computers,OU=DEV Company,DC=dev,DC=local'
            Description = 'Beta location computers'
        }
        @{ 
            Pattern = '^Gamma'
            OU = 'OU=Location Gamma,OU=Computers,OU=DEV Company,DC=dev,DC=local'
            Description = 'Gamma location computers'
        }
        @{ 
            Pattern = '^LAB-'
            OU = 'OU=Lab Computers,OU=Computers,OU=DEV Company,DC=dev,DC=local'
            Description = 'Laboratory computers'
        }
        @{ 
            Pattern = '^KIOSK-'
            OU = 'OU=Kiosk Systems,OU=Computers,OU=DEV Company,DC=dev,DC=local'
            Description = 'Kiosk and public access computers'
        }
        @{ 
            Pattern = '.*'
            OU = 'OU=_Quarantine,OU=Computers,OU=DEV Company,DC=dev,DC=local'
            Description = 'Fallback quarantine for unmatched computers'
        }
    )
}
