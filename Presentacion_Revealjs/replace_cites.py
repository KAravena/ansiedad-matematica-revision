import re
import sys

with open('presentacion_ansiedad_matematica.qmd', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Spans: [@cite]{.microcite} -> [[@cite]]{.microcite}
text = re.sub(r'\[(@[a-zA-Z0-9_]+)\]\{\.microcite\}', r'[[\1]]{.microcite}', text)
text = re.sub(r'\[(@[a-zA-Z0-9_]+);\s*(@[a-zA-Z0-9_]+)\]\{\.microcite\}', r'[[\1; \2]]{.microcite}', text)

# 2. Table rows (Autor/a y año)
text = text.replace('Piccirilli et al. (2023)', '[@piccirilli2023]')
text = text.replace('Aras & Bekdemir (2026)', '[@aras2026]')
text = text.replace('Huang (2026)', '[@huang2026]')
text = text.replace('Kusmaryono et al. (2022)', '[@kusmaryono2022]')
text = text.replace('Shimizu (2025)', '[@shimizu2025]')
text = text.replace('Orbach & Fritz (2022)', '[@orbach2022]')
text = text.replace('Iyamuremye et al. (2022)', '[@iyamuremye2022]')
text = text.replace('Finell et al. (2024)', '[@finell2024]')
text = text.replace('Ma & Sun (2025)', '[@ma2025]')
text = text.replace('Guo & Liao (2022)', '[@guo2022]')

# 3. Grouped citations in the other table
text = text.replace('[@piccirilli2023]; [@finell2024]; [@ma2025]; [@guo2022]', '[@piccirilli2023; @finell2024; @ma2025; @guo2022]')
text = text.replace('[@huang2026]; [@finell2024]; [@guo2022]; [@orbach2022]; [@shimizu2025]; [@aras2026]', '[@huang2026; @finell2024; @guo2022; @orbach2022; @shimizu2025; @aras2026]')
text = text.replace('[@piccirilli2023]; [@huang2026]; [@orbach2022]; [@shimizu2025]; [@iyamuremye2022]; [@ma2025]; [@guo2022]', '[@piccirilli2023; @huang2026; @orbach2022; @shimizu2025; @iyamuremye2022; @ma2025; @guo2022]')
text = text.replace('[@aras2026]; [@huang2026]; [@iyamuremye2022]; [@ma2025]; [@guo2022]', '[@aras2026; @huang2026; @iyamuremye2022; @ma2025; @guo2022]')

# 4. Also any other plaintext citations like the user said:
text = text.replace('Hembree (1990); Ashcraft (2002)', '[@hembree1990nature; @ashcraft2002math]')
text = text.replace('Dowker, Sarkar, y Looi (2016)', '[@dowker2016mathematics]')
text = text.replace('Ashcraft (2002)', '[@ashcraft2002math]')
text = text.replace('Grant y Booth (2009); Munn et al. (2018)', '[@grant2009typology; @munn2018systematic]')
text = text.replace('Page et al. (2021)', '[@page2021prisma]')
text = text.replace('Rethlefsen et al. (2021)', '[@rethlefsen2021prismas]')
text = text.replace("O'Mara-Eves et al. (2015); Marshall y Wallace (2019)", '[@omaraeves2015textmining; @marshall2019automation]')

with open('presentacion_ansiedad_matematica.qmd', 'w', encoding='utf-8') as f:
    f.write(text)

print('Replaced citations successfully.')
