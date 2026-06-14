import re, urllib.parse

html = open('presentacion_ansiedad_matematica.html', encoding='utf-8').read()
m = re.search(r'href="data:text/css,(.*?)"', html)
if m:
    decoded = urllib.parse.unquote(m.group(1))
    with open('decoded.css', 'w', encoding='utf-8') as f:
        f.write(decoded)
    print("Decoded to decoded.css")
    if 'python-scripts-slide' in decoded:
        print("YES! python-scripts-slide IS in the CSS")
    else:
        print("NO! python-scripts-slide IS NOT in the CSS")
else:
    print("Not found")
