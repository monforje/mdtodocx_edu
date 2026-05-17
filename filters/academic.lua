-- filters/academic.lua
-- Pandoc Lua filter: maps markdown elements to Word styles from reference.docx
--
-- Pipeline: MD → pandoc (with this filter) → DOCX → python postprocess → final
--
-- Style mapping:
--   # Title {.centered}  →  center_heading  (centered, bold, 14pt)
--   # 1 Section          →  heading_1       (indent, bold, 14pt, justified)
--   ## 1.1 Subsection    →  heading_2
--   ### 1.1.1 Sub        →  heading_3
--   body paragraphs      →  main_text       (indent, 14pt, 1.5 spacing, justified)
--   bullet lists         →  marker_list
--   numbered lists       →  num_list
--   table captions       →  table_heading   (14pt, letter-spacing)
--   figures (images)     →  main_text + centering (postprocess)
--   image captions       →  image_heading   (12pt, centered)
--   \newpage             →  page break
--   — (em-dash)          →  - (single dash)
--   Page break auto-inserted before every level-1 heading (except first)

-- ============================================================
-- STYLE CONFIG — single source of truth for style names
-- ============================================================

local styles = {
  center_heading = "center_heading",
  heading_1      = "heading_1",
  heading_2      = "heading_2",
  heading_3      = "heading_3",
  main_text      = "main_text",
  marker_list    = "marker_list",
  num_list       = "num_list",
  table_heading  = "table_heading",
  image_heading  = "image_heading",
}

-- Map heading level to style name
local heading_style_map = {
  [1] = styles.heading_1,
  [2] = styles.heading_2,
  [3] = styles.heading_3,
}

-- ============================================================
-- EM-DASH REPLACEMENT: — → -
-- ============================================================

function Str(el)
  el.text = el.text:gsub("—", "-")
  return el
end

-- ============================================================
-- HEADERS — convert ALL to Div with custom-style
-- ============================================================

function Header(el)
  if el.classes:includes("centered") then
    -- # Введение {.centered} → center_heading
    local para = pandoc.Para(el.content)
    return pandoc.Div(para, pandoc.Attr("", {}, {["custom-style"] = styles.center_heading}))
  end

  -- Regular headings → heading_1/2_3 via custom-style
  -- Must use Div because pandoc ignores custom-style on Header elements
  local style = heading_style_map[el.level]
  if style then
    local para = pandoc.Para(el.content)
    -- Store level in class for page-break detection in Pandoc()
    return pandoc.Div(para, pandoc.Attr("", {"heading-" .. el.level}, {["custom-style"] = style}))
  end

  return el
end

-- ============================================================
-- PAGE BREAKS
-- ============================================================

local PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
local PAGE_BREAK_INLINE = '<w:r><w:br w:type="page"/></w:r>'

-- \newpage (LaTeX raw) → OpenXML page break
function RawBlock(el)
  if el.format == "latex" and el.text:match("^\\newpage") then
    return pandoc.RawBlock("openxml", PAGE_BREAK)
  end
  return el
end

-- ::: {.pagebreak} ::: → OpenXML page break
function Div(el)
  if el.classes:includes("pagebreak") then
    return pandoc.RawBlock("openxml", PAGE_BREAK)
  end
  return el
end

-- ============================================================
-- FIGURES (IMAGES) — apply main_text + centering
-- ============================================================

function Figure(el)
  -- Pandoc 3.x wraps images with captions in Figure blocks.
  -- We convert them to styled Divs so postprocess can center them.
  --
  -- Figure {
  --   attr: Attr
  --   caption: Caption (short?, [Block])
  --   content: [Block]  -- usually contains Plain [Image ...]
  -- }

  local result = {}

  -- Image body: wrap in main_text style
  -- Postprocessor detects <w:drawing> and adds centering
  for _, block in ipairs(el.content) do
    if block.t == "Plain" or block.t == "Para" then
      local img_para = pandoc.Para(block.content)
      table.insert(result,
        pandoc.Div(img_para, pandoc.Attr("", {}, {["custom-style"] = styles.main_text}))
      )
    end
  end

  -- Caption: wrap in image_heading style (distinct from table_heading)
  -- Postprocessor converts image_heading → table_heading + centering
  -- Caption has .short (optional [Inline]) and .long ([Block]) fields
  if el.caption and el.caption.long then
    for _, block in ipairs(el.caption.long) do
      if block.t == "Plain" or block.t == "Para" then
        local caption_para = pandoc.Para(block.content)
        table.insert(result,
          pandoc.Div(caption_para, pandoc.Attr("", {}, {["custom-style"] = styles.image_heading}))
        )
      end
    end
  end

  return result
end

-- ============================================================
-- TABLE CAPTIONS — apply table_heading style
-- ============================================================

local function is_table_caption(block)
  -- Pandoc table captions are Para blocks after a Table
  -- They typically start with "Таблица" or "Table"
  if block.t == "Para" then
    local text = pandoc.utils.stringify(block)
    return text:match("^Таблица") or text:match("^Table")
  end
  return false
end

-- ============================================================
-- DOCUMENT-LEVEL: page breaks + style wrapping
-- ============================================================

local function is_level1_heading(block)
  -- Check for Div with heading-1 class (converted from Header)
  if block.t == "Div" then
    for _, cls in ipairs(block.classes) do
      if cls == "heading-1" then
        return true
      end
    end
    -- Also check for center_heading (centered headings are also level-1)
    local cs = block.attributes["custom-style"]
    if cs == styles.center_heading then
      return true
    end
  end
  return false
end

local function is_raw_openxml(block)
  return (
    block.t == "Para" and
    #block.content == 1 and
    block.content[1].t == "RawInline" and
    block.content[1].format == "openxml"
  )
end

function Pandoc(doc)
  local new_blocks = {}
  local seen_header = false
  local prev_was_table = false

  for _, block in ipairs(doc.blocks) do
    if is_level1_heading(block) then
      -- Insert page break before every level-1 heading except the first
      if seen_header then
        table.insert(new_blocks,
          pandoc.Para({pandoc.RawInline("openxml", PAGE_BREAK_INLINE)})
        )
      end
      seen_header = true
      table.insert(new_blocks, block)
      prev_was_table = false

    elseif is_raw_openxml(block) then
      -- Pass through raw OpenXML (page breaks)
      table.insert(new_blocks, block)
      prev_was_table = false

    elseif block.t == "BulletList" or block.t == "OrderedList" then
      -- Pass through lists as-is (pandoc handles numbering, Python fixes styles)
      table.insert(new_blocks, block)
      prev_was_table = false

    elseif is_table_caption(block) then
      -- Wrap table caption in table_heading style
      table.insert(new_blocks,
        pandoc.Div(block, pandoc.Attr("", {}, {["custom-style"] = styles.table_heading}))
      )
      prev_was_table = false

    elseif block.t == "Table" then
      -- Pass through table as-is (postprocess handles formatting)
      table.insert(new_blocks, block)
      -- Insert empty paragraph (native Enter) after table
      -- Use non-breaking space (UTF-8: C2 A0) so pandoc doesn't drop the empty para
      table.insert(new_blocks,
        pandoc.Div(pandoc.Para({pandoc.Str("\194\160")}), pandoc.Attr("", {}, {["custom-style"] = styles.main_text}))
      )
      prev_was_table = true

    elseif block.t == "Para" then
      -- Wrap body paragraphs in main_text style
      -- (table captions already handled above by is_table_caption)
      table.insert(new_blocks,
        pandoc.Div(block, pandoc.Attr("", {}, {["custom-style"] = styles.main_text}))
      )
      prev_was_table = false

    else
      table.insert(new_blocks, block)
      prev_was_table = false
    end
  end
  return pandoc.Pandoc(new_blocks, doc.meta)
end
