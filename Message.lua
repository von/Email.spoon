--- === Email.Message ===
--- Interface to an email message.
--- Wrapper around a dictionary with the following elements:
---  * `from`: a string with an email address
---  * `to`: a list of strings with recipients
---  * `cc`: a list of strings with cc recipients
---  * `bcc`: a list of strings with bcc recipients
---  * `subject`: a string containing the email subject
---  * `content`: a string containing the email content
---  * `attachment`: a list of strings with paths to files

local Message = {}

-- Failed table lookups on the instances should fallback to the class table, to get methods
Message.__index = Message

-- Calls to Message() return Message.new()
setmetatable(Message, {
  __call = function (cls, ...)
    return cls.new(...)
  end,
})

-- Set up logger
Message.log = hs.logger.new("Message")

-- urlencode()
-- Encode a string for inclusion in a URL.
-- Kudos: https://gist.github.com/liukun/f9ce7d6d14fa45fe9b924a3eed5c3d99
--
-- Parameters:
-- * `s`: string to encode
--
-- Reutnes:
-- * Encoded string
local function urlencode(s)
  s = string.gsub(s, "\n", "\r\n")
  s = string.gsub(s, "([^%w])", function (c) return string.format("%%%02X", string.byte(c)) end)
  return s
end

--- Email.Message:debug()
--- Method
--- Enable or disable debugging
---
--- Parameters:
--- * enable: a boolean to indiciate whether debigging should be enabled or disabled
---
--- Returns:
--- * Nothing
function Message:debug(enable)
  if enable then
    self.log.setLogLevel('debug')
    self.log.d("Debugging enabled")
  else
    self.log.d("Disabling debugging")
    self.log.setLogLevel('info')
  end
end

--- Email.Message.new()
--- Constructor
--- Create new Email.Message instance.
---
--- Parameters:
--- * `values` (optional): A table containing initial values
---
--- Returns:
--- * A `Email.Message` instance
function Message.new(values)
  Message.log.d("new() called")
  values = values or {}
  local self = setmetatable(values, Message)
  return self
end

--- Email.Message.fromFile()
--- Constructor
--- Read an email from a file. File format should be zero or more lines of headers,
--- followed by a blank line, followed by zero or more lines of content.
---
--- Parameters:
--- * `path`: string with path to file
---
--- Returns:
--- * `Email.Message` instance, or `nil` on error
function Message.fromFile(path)
  local lines = io.lines(path)
  if not lines then
    self.log.ef("Failed to read %s", path)
    return nil
  end
  local atEOF = false

  local values = {}
  values.to = {}
  values.subject = ""
  values.content = ""

  -- Parse headers
  while true do
    local line = lines()
    -- Blank line == end of headers
    if line == "" then
      break
    end
    -- nil == EOF
    if line == nil then
      atEOF = true
      break
    end
    local field, value = string.match(line, "^(%a+): (.*)$")
    if not field or not value then
      Message.log.ef("Failed to parse header: %s", line)
    else
      field = string.lower(field)
      if field == "to" or field == "cc" or field == "bcc" then
        -- Split commas
        -- Kudos: https://stackoverflow.com/a/19262818/197789
        values[field] = {}
        for addr in string.gmatch(value, '([^,]+)') do
          table.insert(values[field], addr)
        end
      else
        values[field] = value
      end
    end
  end
  if not atEOF then
    -- Parse content
    for line in lines do
      values.content = values.content .. line .. "\n"
    end
  end
  return Message.new(values)
end

--- Message:toURL()
--- Method
--- Return the mail message as a `mailto:` URL.
--- Does not support `from` or `attachment`s.
---
--- Parameters:
--- * None
---
--- Returns:
--- * URL as a string
function Message:toURL()
  local url = "mailto:"
  if self.to then
    url = url .. table.concat(self.to, ",")
  end
  local queries = {}
  if self.subject then
    table.insert(queries, "subject=" .. urlencode(self.subject))
  end
  if self.cc then
    table.insert(queries, "cc=" .. table.concat(self.cc, ","))
  end
  if self.bcc then
    table.insert(queries, "bcc=" .. table.concat(self.bcc, ","))
  end
  if self.content then
    table.insert(queries, "body=" .. urlencode(self.content))
  end
  if #queries > 0 then
    url = url .. "?" .. table.concat(queries, "&")
  end
  return url
end

return Message
-- vim: foldmethod=marker: --
