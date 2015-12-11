#empty links
$global:links = @()

$orders = @{0="Common"; 1="Other"}

$Path = "H:\Shortcuts\Bookmarks"
cd $Path

$head = Get-Content C:\Users\M136815\Desktop\Powershell\head.txt

function parseURLFile ($item, $Path) {
    $dir = Split-Path (Split-Path $item -Parent) -Leaf
    $linktext = $item.Replace(".URL", "").Replace($Path,"").Replace("\","")
    $content = Get-Content $item | select-string "URL="
    $url = $content -replace "URL=", ""

    $object = New-Object System.Object
    $object | Add-Member -Name 'Link' -type NoteProperty -Value $url
    $object | Add-Member -Name 'LinkText' -type NoteProperty -Value $linktext
    $object | Add-Member -Name 'Container' -type NoteProperty -Value $dir

    $global:links += $object
}


function recurseDirForURLFiles($items) {
    ForEach ($item in $items) { 
        If(Test-Path $item.FullName -pathtype container){
            $subItems = Get-ChildItem $item.FullName | select FullName
            recurseDirForURLFiles $subItems

        } Else {

            parseURLFile $item.FullName $Path 
        }
    }
}

$items = Get-ChildItem $Path | select FullName
recurseDirForURLFiles $items

$containers = New-Object System.Collections.ArrayList

$links | `
Foreach-Object {
##Write-Debug $_.Link
#Write-Debug $_.LinkText
If($containers -notcontains $_.Container) {
    $containers += $_.Container
}
}

$orderedContainers = New-Object System.Collections.ArrayList

For ($i=0; $i -le $orders.length; $i++) {
    $orderedContainers.Add($orders.Get_Item($i))
    }





$body ='<h2>MK Work Links</h2>'

$body +='<div class="theGroups">'

Foreach($container in $orderedContainers){
   #Write-Debug "create div for container $container"
   $body += '<div class="group">'
   $body += '<div class="title">' + $container + '</div>'
   $body += '<div class="theLinks">'
   $linksForContainer = $links | where-object {$_.Container -eq $container}
   Foreach($linkForContainer in $linksForContainer){
   $linkText = ($linkForContainer.LinkText).Replace($container,"")
   $link = $linkForContainer.Link
   Write-Host "LinkText:$linkText" 
   #Write-Debug "url:" $linkForContainer.Link

   $body += '<div class="notUpdated"><a href="' + $linkForContainer.Link + '" style="color: #000;" target="_self">' + $linkText + '</a></div>'
   }
   $body += '</div>'
   $body += '</div>'

}

$body += '<div class="belowSpace"></div>'
$body += '</div>'

ConvertTo-HTML -head $head -body $body | Out-File H:\bookmark7.html
##Write-Debug $body