function Sort-RunPropertyOrder {
    <#
    .SYNOPSIS
        Reorders the elements inside a run-properties block into the sequence
        the OOXML schema requires.

    .DESCRIPTION
        The children of <w:rPr> are a sequence, not a set: the schema fixes their
        order (rFonts, then b, then i, ... then color, spacing, sz, and so on).
        Word reads them in any order, so a scrambled rPr renders correctly and
        nothing appears to be wrong. Stricter readers - XML validators, pandoc,
        LibreOffice, a submission check that validates rather than opens - are
        entitled to reject it.

        Successive regex insertions are what scramble it. A function that writes
        rFonts just after <w:rPr>, then writes sz just after <w:rPr> as well,
        leaves sz sitting in front of the rFonts it wrote first. Sorting
        afterwards is simpler than making every insertion position-aware.

        Only reorders. Nothing is added, removed, or edited, so it is safe to run
        over any block and is idempotent.

        Elements the schema list does not name (w14:* extensions such as
        w14:ligatures, and anything from a later spec revision) keep their
        relative order and are placed after the known elements. rPrChange is
        forced last, which is where the schema puts it.

    .PARAMETER Block
        A style block, or any XML fragment containing a <w:rPr>. Every rPr in the
        fragment is sorted.

    .EXAMPLE
        $block = Sort-RunPropertyOrder -Block $block
    #>
    param(
        [string]$Block
    )

    # CT_RPr child sequence, ECMA-376 Part 1, 17.3.2.28.
    $order = @(
        'rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike',
        'outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden',
        'color','spacing','w','kern','position','sz','szCs','highlight','u','effect',
        'bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout',
        'specVanish','oMath'
    )
    $rank = @{}
    for ($i = 0; $i -lt $order.Count; $i++) { $rank[$order[$i]] = $i }

    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $open  = $m.Groups[1].Value
        $inner = $m.Groups[2].Value
        $close = $m.Groups[3].Value
        if ([string]::IsNullOrWhiteSpace($inner)) { return $m.Value }

        # Top-level children: self-closing, or an open/close pair. rPrChange is
        # the only one that nests, and the lazy match stops at its own close tag.
        $childPattern = '(?s)<(\w+):(\w+)(?:\s[^>]*?)?/>|<(\w+):(\w+)(?:\s[^>]*?)?>.*?</\1:\2>'
        $children = [regex]::Matches($inner, $childPattern)
        if ($children.Count -eq 0) { return $m.Value }

        # Anything the pattern did not consume (stray text) means this rPr is not
        # shaped the way we expect, so leave it exactly as it was.
        $consumed = ($children | ForEach-Object { $_.Length } | Measure-Object -Sum).Sum
        if ($consumed -ne $inner.Length) { return $m.Value }

        $items = @()
        $seq = 0
        foreach ($c in $children) {
            $local = if ($c.Groups[2].Success) { $c.Groups[2].Value } else { $c.Groups[4].Value }
            $prefix = if ($c.Groups[1].Success) { $c.Groups[1].Value } else { $c.Groups[3].Value }
            if ($local -eq 'rPrChange') {
                $key = [int]::MaxValue          # always last
            } elseif ($prefix -eq 'w' -and $rank.ContainsKey($local)) {
                $key = $rank[$local]
            } else {
                # w14:* and anything unrecognised: after the known elements,
                # before rPrChange, in the order they were already in.
                $key = $order.Count + 1
            }
            $items += [pscustomobject]@{ Key = $key; Seq = $seq; Xml = $c.Value }
            $seq++
        }

        $sorted = $items | Sort-Object Key, Seq
        return $open + (($sorted | ForEach-Object { $_.Xml }) -join '') + $close
    }

    return [regex]::Replace($Block, '(?s)(<w:rPr(?:\s[^>]*?)?>)(.*?)(</w:rPr>)', $evaluator)
}
