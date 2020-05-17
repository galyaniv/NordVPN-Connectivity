Function Average($speed_list)
{
    $sum_speed = 0;
    foreach($s in $speed_list){
        $sum_speed += $s
    }
    return ([decimal]($sum_speed) / [decimal]($speed_list.Length));
}

$data_file_path = [System.IO.Path]::Combine($($psISE.CurrentFile.FullPath | Split-Path -Parent) ,"data.txt")

$data_file_content = Import-Csv -Path $data_file_path -Header "ServerName","Date","Time","Download Speed","Upload Speed","Ping OK","Ping Lost"

$dict = New-Object 'system.collections.generic.dictionary[string,decimal]'

[System.Decimal[]]$s_list = @()

foreach ($item in $data_file_content | Group-Object ServerName | sort Count  -Descending| Group-Object ServerName| %{$_.Group}){
    [System.Decimal[]]$s_list += $item | %{$_.Group."Download Speed"}
    $Average = Average($s_list)
    $server_name = $item | %{$_.Name}
    $dict[$server_name] = [math]::Round($Average, 2)

}
$dict.GetEnumerator() | sort -Property Value -Descending 