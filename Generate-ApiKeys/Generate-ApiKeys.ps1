param(
    [Parameter()]
    [ValidateRange(1,32)]
    [int]$Count = 10,

    [Parameter()]
    [switch]$DontRandomizeDashes
)

if ($DontRandomizeDashes)
{
    [int[]]$t=5,13,24,33,45,57,62,71,82,88,91,12,13,14,15;
    [bool]$tSet=$true;
    [int]$tVal=10;
} else {
    [int[]]$t=1,2,3,4,5,6,7,8,9,10,11,12,13,14,15;
    [bool]$tSet=$false;
    [int]$tVal=0;
}
$i=1;

do{
    $bytes=[System.Text.Encoding]::Unicode.GetBytes((new-guid).ToString());
    $e=[Convert]::ToBase64String($bytes);
    $salt = "ThisIsADumbSaltValue" + [DateTime]::UtcNow.ToString("yyyyMMdd_HHmmss.ffffff_UTC");
    $str=[IO.MemoryStream]::new([byte[]][char[]]($e + $salt));
    $hash=(Get-FileHash -InputStream $str -Algorithm SHA384).Hash.ToLower()
    [int16]$l = $(Get-Random -Maximum 12 -Minimum 5)
    if ($tSet -eq $false)
    {
        $t[$tVal]=5;
        while($l -lt $hash.ToString().Length-5)
        {
            $tVal+=1;
            #$tVal;
            $min = $(if(($l + 5) -ge $hash.Length-3) {$l+1} else {$l + 5});
            if ($tval % 2 -eq 0)
            {
                if ($mon-$l -lt 11) {$min = $l+11}
                if ($min -gt $hash.Length) {$min = $hash.Length - $(Get-Random -Minimum 8 -Maximum 10)}
            }  
            $max = $(if(($l + 17) -ge $hash.Length-3) {$hash.Length} else {$l+17});
            if ($min -ge $max) {$min = $min - 5}
            $l = Get-Random -Minimum $min -Maximum $max;
            $t[$tVal]=$l;
        }
        $tSet=$true;
    }
    for($w=0;$w -lt $tVal ;$w++)
    {
        #$t.Count
        #$t[$w];
        if (($hash.Substring($t[$w]-1,1) -eq '-') -or ($hash.Substring($t[$w]-2,1) -eq '-') -or ($hash.Substring($t[$w]-3,1) -eq '-') -or ($hash.Substring($t[$w]-4,1) -eq '-') )
        {

        } else {
            $hash=$hash.Remove($t[$w],1).Insert($t[$w],"-");            
        }
    }

    $hash
    $i+=1;
    #exit
}until($i -ge $count)

#$t