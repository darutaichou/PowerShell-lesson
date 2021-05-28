$Args[0] -match "_‹Î–±•\_(?<month>.*?)ŒŽ" | Out-Null

$Matches.month

resolve-path $Args[0]