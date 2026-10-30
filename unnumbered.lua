-- Lua filter for `make docx` / `make odt`. Must run BEFORE --citeproc.
--
-- 1. Unnumbered headings (\section*, \subsection*): the heading styles in
--    reference.docx are numbered automatically, so Word would also count
--    "Instructions" and "References". For docx, such headings are written with
--    the normal heading style but with numbering switched off (numId 0), so they
--    keep the font and size of all other headings.
-- 2. References: pandoc ignores \printbibliography and appends the reference
--    list at the very end (after the appendix). An empty Div with id "refs"
--    after the "References" heading tells citeproc where to put the list.

local function xml_escape(s)
  return (s:gsub('&', '&amp;'):gsub('<', '&lt;'):gsub('>', '&gt;'))
end

local function unnumbered_docx(el)
  local text = xml_escape(pandoc.utils.stringify(el.content))
  local level = math.min(el.level, 9)
  return pandoc.RawBlock('openxml',
    '<w:p><w:pPr><w:pStyle w:val="Heading' .. level .. '"/>' ..
    '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>' ..
    '<w:r><w:t xml:space="preserve">' .. text .. '</w:t></w:r></w:p>')
end

function Pandoc(doc)
  local blocks = pandoc.List()
  local refs_placed = false
  for _, b in ipairs(doc.blocks) do
    if b.t == 'Header' and b.classes:includes('unnumbered') then
      if FORMAT:match('docx') then
        blocks:insert(unnumbered_docx(b))
      else
        blocks:insert(b)
      end
      if not refs_placed and pandoc.utils.stringify(b.content) == 'References' then
        blocks:insert(pandoc.Div({}, pandoc.Attr('refs')))
        refs_placed = true
      end
    else
      blocks:insert(b)
    end
  end
  doc.blocks = blocks
  return doc
end
