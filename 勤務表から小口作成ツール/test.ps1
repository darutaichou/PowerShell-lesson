$Args[0] -match "_�Ζ��\_(?<month>.*?)��" | Out-Null

$Matches.month

resolve-path $Args[0]