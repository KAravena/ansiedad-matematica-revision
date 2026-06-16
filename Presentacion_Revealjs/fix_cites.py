with open('presentacion_ansiedad_matematica.qmd', 'r', encoding='cp1252') as f:
    text = f.read()

# Fix Citations block 1
text = text.replace(
    '<span class="microcite">(Hembree, 1990; Ashcraft, 2002)</span>',
    '<span class="microcite">[@hembree1990nature; @ashcraft2002math]</span>'
)
text = text.replace(
    '<span class="microcite">(Dowker et al., 2016)</span>',
    '<span class="microcite">[@dowker2016mathematics]</span>'
)
text = text.replace(
    '<span class="microcite">(Ashcraft, 2002; Dowker et al., 2016)</span>',
    '<span class="microcite">[@ashcraft2002math; @dowker2016mathematics]</span>'
)

# Fix Citations block 2
text = text.replace(
    '<span class="microcite">(Hembree, 1990; Dowker et al., 2016)</span>',
    '<span class="microcite">[@hembree1990nature; @dowker2016mathematics]</span>'
)
text = text.replace(
    '<span class="microcite">(Ashcraft, 2002)</span>',
    '<span class="microcite">[@ashcraft2002math]</span>'
)

# Fix Citations block 3 (Grant & Booth)
text = text.replace(
    '<span class="microcite">(Grant & Booth, 2009; Munn et al., 2018)</span>',
    '<span class="microcite">[@grant2009typology; @munn2018systematic]</span>'
)
text = text.replace(
    '<span class="microcite">(Page et al., 2021)</span>',
    '<span class="microcite">[@page2021prisma]</span>'
)
text = text.replace(
    '<span class="microcite">(Munn et al., 2018)</span>',
    '<span class="microcite">[@munn2018systematic]</span>'
)

# Strip BOM if it magically reappears
if text.startswith('\ufeff'):
    text = text[1:]

with open('presentacion_ansiedad_matematica.qmd', 'w', encoding='utf-8') as f:
    f.write(text)
