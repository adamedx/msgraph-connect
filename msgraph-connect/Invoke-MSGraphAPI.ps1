using assembly ".\msgraph-connect.dll"

# Copyright 2020, Adam Edwards
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

set-strictmode -version 2

function Invoke-MSGraphAPI {
    [cmdletbinding(positionalbinding=$false)]
    param(
        [parameter(position=0, mandatory=$true)]
        [string] $Uri,

        [string] $Method = 'GET',

        [parameter(parametersetname='simple')]
        [string] $Version = '1.0',

        [parameter(parametersetname='simple')]
        [Uri] $GraphUri = 'https://graph.microsoft.com',

        [parameter(parametersetname='simple')]
        [Uri] $LoginUri = 'https://login.microsoftonline.com',

        [parameter(parametersetname='simple')]
        [string[]] $Permissions = @('User.Read'),

        [parameter(parametersetname='simple')]
        [string] $AppId,

        [switch] $FullResponse,

        [switch] $RawContent,

        [parameter(parametersetname='connection', mandatory=$true)]
        [msgraph_connect.GraphConnection] $Connection
    )

    $targetAppId = if ( $AppId ) {
        [Guid] $AppId
    }

    $graphConnection = if ( $Connection ) {
        $Connection
    } else {
        [msgraph_connect.GraphConnection]::new($Permissions, $GraphUri, $LoginUri, $Version, $targetAppId)
    }

    write-host $uri, $method

    $graphConnection | fl * | out-host
    $response = $graphConnection.InvokeRequest($Uri, $Method)

    $global:mycon = $graphConnection
    if ( $FullResponse.IsPresent ) {
        $response
    } else {
        $content = $response.Content
        if ( $RawContent.IsPresent ) {
            $content
        } else {
            $content | convertfrom-json
        }
    }
}

function Connect-MSGraphAPI {
    [cmdletbinding(positionalbinding=0)]
    param(
        [string] $Version,

        [Uri] $GraphUri = 'https://graph.microsoft.com',

        [Uri] $LoginUri = 'https://login.microsoftonline.com',

        [string[]] $Permissions = @('User.Read'),

        [string] $AppId,

        [switch] $NoConnect
    )

    $connection = [msgraph_connect.GraphConnection]::new(
        $Permissions,
        $GraphUri,
        $LoginUri,
        $Version,
        $AppId)

    if ( ! $NoConnect.IsPresent ) {
        $connection.Connect()
    }

    $connection
}
