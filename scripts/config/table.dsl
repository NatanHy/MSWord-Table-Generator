!bold("Variable") | !span(!bold("Variable influence on process"), 3) | !span(!bold("Process influence on variables"), 3) !newline
!bold("Variable") | !bold("Influence present? (Yes/No Description)") | !bold("Time period/Climate domain") | 
!bold("Handling of influence (How/If not - Why)") | !bold("Influence present? (Yes/No Description)") | 
!bold("Time period/Climate domain") | !bold("Handling of influence (How/If not - Why)") !newline

foreach(!variables) as $var {
    foreach(!time_period) as $time {
        !description($var)
        foreach(!influence) as $influence {
            [$var][$influence]["Influence present?"]["Yes/No"] + [$var][$influence]["Influence present?"]["Description"] | 
            $time | 
            [$var][$influence][$time]["Rationale"]
        }
        !newline
    }
    !force_cutoff
}