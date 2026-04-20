local system = require("pandoc.system")

local mermaid_dir = os.getenv("MARKDOWN_DOCX_MERMAID_DIR")
local mermaid_format = os.getenv("MARKDOWN_DOCX_MERMAID_FORMAT") or "png"
local mermaid_scale = os.getenv("MARKDOWN_DOCX_MERMAID_SCALE") or "2"
local mermaid_width = os.getenv("MARKDOWN_DOCX_MERMAID_WIDTH") or "5.5in"
local mermaid_cmd = os.getenv("MERMAID_MMDC") or "mmdc"
local mermaid_puppeteer_config = os.getenv("MARKDOWN_DOCX_MERMAID_PUPPETEER_CONFIG")
  or os.getenv("MERMAID_PUPPETEER_CONFIG")
local ensured_mermaid_dir = false
local mermaid_counter = 0
local current_header = nil

local function file_exists(path)
  local handle = io.open(path, "rb")
  if handle then
    handle:close()
    return true
  end

  return false
end

local function write_file(path, contents)
  local handle, err = io.open(path, "wb")
  if not handle then
    error("Unable to write Mermaid source file: " .. tostring(err))
  end

  handle:write(contents)
  handle:close()
end

local function read_file(path)
  if not path or path == "" then
    return ""
  end

  local handle = io.open(path, "rb")
  if not handle then
    return ""
  end

  local contents = handle:read("*a") or ""
  handle:close()
  return contents
end

local function ensure_mermaid_dir()
  if ensured_mermaid_dir then
    return
  end

  if not mermaid_dir or mermaid_dir == "" then
    error("Mermaid rendering requires MARKDOWN_DOCX_MERMAID_DIR to be set.")
  end

  system.make_directory(mermaid_dir, true)
  ensured_mermaid_dir = true
end

local function text_to_inlines(text)
  local inlines = pandoc.List()

  for word, spacing in text:gmatch("(%S+)(%s*)") do
    inlines:insert(pandoc.Str(word))
    if spacing ~= "" then
      inlines:insert(pandoc.Space())
    end
  end

  if #inlines == 0 then
    inlines:insert(pandoc.Str("Mermaid"))
    inlines:insert(pandoc.Space())
    inlines:insert(pandoc.Str("diagram"))
  end

  return inlines
end

local function render_mermaid(source_path, output_path)
  local args = {
    "-i", source_path,
    "-o", output_path,
    "--backgroundColor", "transparent",
  }

  if mermaid_format == "png" then
    table.insert(args, "--scale")
    table.insert(args, mermaid_scale)
  end

  if mermaid_puppeteer_config and mermaid_puppeteer_config ~= "" then
    table.insert(args, "-p")
    table.insert(args, mermaid_puppeteer_config)
  end

  local ok, result = pcall(pandoc.pipe, mermaid_cmd, args, "")
  if not ok then
    error(
      "Failed to render Mermaid diagram with "
        .. mermaid_cmd
        .. ". Install Mermaid CLI (`npm install -g @mermaid-js/mermaid-cli`) "
        .. "or set MERMAID_MMDC to the executable path.\n"
        .. tostring(result)
    )
  end
end

local function image_attr(block)
  local attrs = {}

  if block.attributes["width"] then
    attrs["width"] = block.attributes["width"]
  elseif mermaid_width ~= "" then
    attrs["width"] = mermaid_width
  end

  if block.attributes["height"] then
    attrs["height"] = block.attributes["height"]
  end

  return pandoc.Attr(block.identifier, {}, attrs)
end

function CodeBlock(block)
  if not block.classes:includes("mermaid") then
    return nil
  end

  mermaid_counter = mermaid_counter + 1
  ensure_mermaid_dir()

  local cache_key = pandoc.sha1(table.concat({
    block.text,
    mermaid_format,
    mermaid_scale,
    mermaid_cmd,
    mermaid_puppeteer_config or "",
    read_file(mermaid_puppeteer_config),
  }, "|"))
  local source_path = mermaid_dir .. "/" .. cache_key .. ".mmd"
  local output_path = mermaid_dir .. "/" .. cache_key .. "." .. mermaid_format

  if not file_exists(output_path) then
    write_file(source_path, block.text)
    render_mermaid(source_path, output_path)
    os.remove(source_path)
  end

  local caption_label = block.attributes["caption"]
    or block.attributes["title"]
    or current_header
    or "Mermaid diagram"
  local image_para = pandoc.Para({
    pandoc.Image(text_to_inlines("Mermaid diagram"), output_path, "", image_attr(block)),
  })
  local caption_para = pandoc.Para(
    text_to_inlines("Figure " .. mermaid_counter .. ". " .. caption_label),
    pandoc.Attr("", {}, { ["custom-style"] = "Image Caption" })
  )

  return { image_para, caption_para }
end

function Header(block)
  current_header = pandoc.utils.stringify(block.content)
  return nil
end
