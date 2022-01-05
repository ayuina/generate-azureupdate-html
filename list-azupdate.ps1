<#
    .SYNOPSIS
    Retrieve data from Azure Update (https://azure.microsoft.com/ja-jp/updates/)
    
    .DESCRIPTION
    This script download rss feed, filter by date, and output as html format.
    
    .PARAMETER from
    Filter option from date, this parameter require only date part and ignore time part

    .PARAMETER to
    Filter option to date, this parameter require only date part and ignore time part
    
#>

[CmdletBinding()]
param (

    [parameter(Mandatory=$true)]
    [System.DateTimeOffset]$from,
    [parameter(Mandatory=$true)]
    [System.DateTimeOffset]$to 
)

function Main 
{

    [System.DateTimeOffset]$startoffromdate = $from.Date
    [System.DateTimeOffset]$endoftodate = $to.Date.AddDays(1).AddMilliseconds(-1)
    Write-Verbose "retrieving from ${startoffromdate} to ${endoftodate}"

    $feed = GetRssFeed

    $target = $feed.CreateNavigator().Select('//item') `
    | Where-Object {
        $pubdate = [System.DateTimeOffset]::Parse( $_.SelectSingleNode('pubDate').Value )
        $between = (($startoffromdate -le $pubdate ) -and ($pubdate -le $endoftodate ))
        Write-Output $between
        Write-Verbose ("$pubdate " + ($between ? "is " : "is not ") + "between $startoffromdate and $endoftodate")
    } `
    | ForEach-Object { 
        [PSCustomObject]@{ 
            title = $_.SelectSingleNode('title').Value;
            link = $_.SelectSingleNode('link').Value;
            pubDate = $_.SelectSingleNode('pubDate').Value ;
            description = $_.SelectSingleNode('description').Value }
    } 

    $pagetitle = ("{0:yyyyMMdd}-{1:yyyyMMdd} Azure Updates" -f $startoffromdate, $endoftodate)
    $timestamp = "{0:yyyy/MM/dd HH:mm:ss}" -f [datetime]::Now
    $output = $target | ConvertTo-Html `
        -Title $pagetitle `
        -PreContent "<h1>$pagetitle</h1> generated at ${timestamp} <hr/>" `
        -Property ( 
            @{name = 'LastUpdate'; expr={[System.Environment]::NewLine + $_.pubDate  } },
            @{name = 'Title'; expr={[System.Environment]::NewLine + "<a href='$($_.link)'>$($_.title)</a>"  } },  
            @{name = 'Description'; expr={[System.Environment]::NewLine + $_.description  } }  
        )

    $htmlfile = ".\${pagetitle}.htm"
    Write-Verbose "output as html file ${htmlfile}"
    Add-Type -AssemblyName System.Web
    [System.Web.HttpUtility]::HtmlDecode($output) | Out-File -FilePath $htmlfile

}

function GetRssFeed()
{
    $feedurl = 'https://azurecomcdn.azureedge.net/ja-jp/updates/feed/'
    try {
        $res = Invoke-WebRequest -Method Get -Uri $feedurl
    }
    catch {
        throw $_
    }
    Write-Verbose "successfully retrieved rss feed"
    return [xml]($res.Content)

}


Main




