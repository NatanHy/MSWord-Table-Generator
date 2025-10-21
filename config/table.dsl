!format("style='Table Grid', width=Cm(16)")
!style("bold=True, name='Times New Roman', size=Pt(10)")

"Variable" | !span("Variable influence on process", 3) | !span("Process influence on variables", 3) !newline
"Variable" | "Influence present? (Yes/No Description)" | "Conditional domain" | 
"Handling of influence (How/If not - Why)" | "Influence present? (Yes/No Description)" | 
"Time period/Climate domain" | "Handling of influence (How/If not - Why)" !newline

!style("name='Times New Roman', size=Pt(10)")
foreach(!variables) as $var {
    foreach(!time_period) as $time {
        !description($var)
        foreach(!influence) as $influence {
            if (!equals([$var][$influence]["Influence present?"]["Yes/No"], "No")) {
                "No" | "N/A" | "N/A"
            } else {
                [$var][$influence]["Influence present?"]["Yes/No"] + "\n" + [$var][$influence]["Influence present?"]["Description"] | 
                $time | 
                [$var][$influence][$time]["Rationale"]
            }
        }
        !newline
    }
    !force_cutoff
}