Import-Module SQLServer

# Create SQL connection string
    $connectionString = "Server=$discoSQL;Database=Disco;User Id=$(($Secret:discorw).UserName);Password=$(($Secret:discorw).GetNetworkCredential().Password);TrustServerCertificate=True;"

try {
    # Connect to SQL Server and retrieve device data
    Write-Information "Connecting to database..."
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = $ConnectionString
    $SqlConnection.Open()

    # Query to get devices
    $Query = "SELECT SerialNumber, AssetNumber, ComputerName FROM Devices"
    $SqlCommand = New-Object System.Data.SqlClient.SqlCommand
    $SqlCommand.Connection = $SqlConnection
    $SqlCommand.CommandText = $Query

    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCommand
    $DataSet = New-Object System.Data.DataSet
    $SqlAdapter.Fill($DataSet)

    $Devices = $DataSet.Tables[0]
    Write-Information "Retrieved $($Devices.Rows.Count) devices from database"

    # Process each device
    foreach ($Device in $Devices.Rows) {
        $SerialNumber = $Device.SerialNumber
        $CurrentAssetNumber = $Device.AssetNumber
        $DeviceName = $Device."Device Name"

        Write-Verbose "Processing device: $DeviceName (Serial: $SerialNumber)"

        try {
            # Query Snipe-IT API for the serial number
            $Headers = @{
                "Authorization" = "Bearer $($Secret:icttoolssnipero)"
                "Accept" = "application/json"
                "Content-Type" = "application/json"
            }

            $SnipeITAPIUrl = "$SnipeITURL/api/v1/hardware/byserial/$SerialNumber"
            $Response = Invoke-RestMethod -Uri $SnipeITAPIUrl -Headers $Headers -Method GET

            if ($Response.rows -and $Response.rows.Count -gt 0) {
                # Asset found in Snipe-IT
                $SnipeITAsset = $Response.rows[0]
                $SnipeITAssetNumber = $SnipeITAsset.asset_tag

                Write-Verbose "Found asset in Snipe-IT: $SnipeITAssetNumber"

                # Compare asset numbers
                if ($CurrentAssetNumber -ne $SnipeITAssetNumber) {
                    Write-Information "Asset number mismatch! Database: $CurrentAssetNumber, Snipe-IT: $SnipeITAssetNumber"

                    # Update database with correct asset number
                    $UpdateQuery = "UPDATE Devices SET AssetNumber = @NewAssetNumber WHERE SerialNumber = @SerialNumber"
                    $UpdateCommand = New-Object System.Data.SqlClient.SqlCommand($UpdateQuery, $SqlConnection)
                    $UpdateCommand.Parameters.AddWithValue("@NewAssetNumber", $SnipeITAssetNumber) | Out-Null
                    $UpdateCommand.Parameters.AddWithValue("@SerialNumber", $SerialNumber) | Out-Null

                    $RowsAffected = $UpdateCommand.ExecuteNonQuery()

                    if ($RowsAffected -gt 0) {
                        Write-Information "Successfully updated asset number for $SerialNumber"
                    } else {
                        Write-Information "Failed to update asset number for $SerialNumber"
                    }

                    # Clean up command object
                    $UpdateCommand.Dispose()
                } else {
                    Write-Verbose "Asset numbers match - no update needed"
                }
            } else {
                Write-Information "Asset not found in Snipe-IT for serial number: $SerialNumber"
            }
        }
        catch {
            Write-Information "Error processing device $SerialNumber : $($_.Exception.Message)"
        }

        # Add small delay to avoid overwhelming the API
        Start-Sleep -Milliseconds 500
    }
}
catch {
    Write-Information "Database connection error: $($_.Exception.Message)"
}
finally {
    # Clean up database connection
    if ($SqlConnection.State -eq 'Open') {
        $SqlConnection.Close()
        Write-Information "Database connection closed"
    }
}

Write-Information "Script completed"