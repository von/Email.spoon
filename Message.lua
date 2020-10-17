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
  local values = {}
  -- Parse headers
  while true do
    local line = lines()
    if line == "" then
      break
    end
    local field, value = string.match(line, "^(%a+): (.*)$")
    if not field or not value then
      self.log.ef("Failed to parse header: %s", line)
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
  -- Parse content
  values.content = ""
  for line in lines do
    values.content = email.content .. line
  end
  return Email.new(values)
end

return Message
-- vim: foldmethod=marker: --
