import codecs

css_to_append = """
/* =========================================================
   Citas visibles homogéneas - RevealJS
   ========================================================= */

.reveal .microcite,
.reveal .slide-citations,
.reveal .method-citation,
.reveal .citation-note,
.reveal .refs-note {
  color: #6f8fae !important;
  font-weight: 500 !important;
}

.reveal .microcite {
  display: block !important;
  margin-top: 0.22rem !important;
  font-size: 0.72rem !important;
  line-height: 1.18 !important;
}

.reveal .slide-citations,
.reveal .method-citation,
.reveal .citation-note {
  font-size: 0.88rem !important;
  line-height: 1.25 !important;
  text-align: center !important;
}

/* Cualquier enlace de cita generado por Pandoc/Citeproc */
.reveal a[href^="#ref-"],
.reveal a[role="doc-biblioref"],
.reveal .citation a,
.reveal .slide-citations a,
.reveal .method-citation a,
.reveal .microcite a,
.reveal .citation-note a,
.reveal table a[href^="#ref-"],
.reveal td a[href^="#ref-"] {
  color: #6f8fae !important;
  text-decoration: none !important;
  font-weight: 500 !important;
}

/* Evitar verde intenso en citas dentro de tablas o matrices */
.reveal table a[href^="#ref-"]:hover,
.reveal td a[href^="#ref-"]:hover,
.reveal a[href^="#ref-"]:hover,
.reveal a[role="doc-biblioref"]:hover {
  color: #6f8fae !important;
  text-decoration: underline !important;
}

/* =========================================================
   Popup de citas / bibliografía al pasar el mouse
   ========================================================= */

.tippy-box,
.tippy-box[data-theme~="quarto"],
.tippy-box[data-theme~="light-border"] {
  max-width: 520px !important;
  font-size: 0.72rem !important;
  line-height: 1.2 !important;
  color: #2f3b45 !important;
  background: #ffffff !important;
  border: 1px solid rgba(95, 111, 124, 0.35) !important;
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.16) !important;
}

.tippy-content {
  padding: 0.55rem 0.65rem !important;
  font-size: 0.72rem !important;
  line-height: 1.2 !important;
}

.tippy-content .csl-entry,
.tippy-content p,
.tippy-content div {
  font-size: 0.72rem !important;
  line-height: 1.2 !important;
  margin-bottom: 0.35rem !important;
}

.tippy-content a {
  color: #6f8fae !important;
  text-decoration: none !important;
}
"""

with codecs.open('libs/ansiedad-overrides.scss', 'a', 'utf-8') as f:
    f.write('\n' + css_to_append)
