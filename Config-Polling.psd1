@{
    # Log configuration
    LogDir = '.\Logs'
    MaxLogSizeMB = 10
    EnableTranscript = $true
    
    # Polling configuration
    PollIntervalMinutes = 3               # How often to run the polling cycle
    LookbackMinutes = 10                  # How far back to search for new computers
    MaxProcessedHistoryDays = 30          # How long to keep processed status in AD (extensionAttribute7)
    
    # Error handling and retry
    RetryFailedMoves = $true              # Retry failed computer moves
    MaxRetryAttempts = 3                  # Maximum retry attempts per computer
    RetryIntervalMinutes = 5              # Wait time between retries
    
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
    
    # Advanced settings
    MaxConcurrentDCQueries = 5            # Limit concurrent DC queries to prevent overload
    DCQueryTimeoutSeconds = 30            # Timeout for individual DC queries
    SkipDCsWithErrors = $true            # Skip DCs that are unreachable rather than failing
}