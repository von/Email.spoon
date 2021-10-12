--- === Email.App ===
---
--- An abstract base class for email application interfaces

local App = {}

-- Failed table lookups on the instances should fallback to the class table, to get methods
App.__index = App

setmetatable(App, {
  -- Calls to App() and subclasses return <class>.new()
  __call = function (cls, ...)
    return cls.new(...)
  end,
})

local Message = dofile(hs.spoons.resourcePath("Message.lua"))

-- Set up logger
App.log = hs.logger.new("Email.App")

--- Email.App:debug()
--- Method
--- Enable or disable debugging
---
--- Parameters:
--- * enable: a boolean to indiciate whether debigging should be enabled or disabled
---
--- Returns:
--- * Nothing
function App:debug(enable)
  if enable then
    self.log.setLogLevel('debug')
    self.log.d("Debugging enabled")
  else
    self.log.d("Disabling debugging")
    self.log.setLogLevel('info')
  end
end

-- Email.App.new()
-- Constructor
-- Create a new instanace of Email.App
-- Not meant to be called directly, but via a subclass.
--
-- Parameters:
-- * None
--
-- Returns:
-- * Email.App instance.
function App.new()
  App.log.d("new() called")
  local self = setmetatable({}, App)
  self.useHTML = false
  return self
end

--- Email.App:escapeApplescriptString()
--- Method
--- Escape quotes and backslashes in given string for inclusion in Applescript.
---
--- Parameters:
--- * s: String to escape.
---
--- Returns:
--- * A string appropriate escaped.
function App:escapeApplescriptString(s)
  s = string.gsub(s, [[\]], [[\\]])
  s = string.gsub(s, [["]], [[\"]])
  return s
end

--- Email.App:executeApplescript()
--- Method
--- Execute given Applescript.
---
--- Parameters:
--- * script: Applescript as a string
---
--- Returns:
--- * true on success, false on error
function App:executeApplescript(script)
  self.log.d("Executing script:" .. script)
  local success, obj, output = hs.osascript.applescript(script)
  if not success then
    hs.alert("Mail composition script failed")
    self.log.e(output.OSAScriptErrorMessageKey)
    self.log.e("Script: ")
    self.log.e(script)
    return false
  end
  self.log.d("Success.")
  return true
end

-- Email.App:compose()
-- Method
-- Given an Email.Message instance, create new email comosition
-- with its contents. Must be overridden by subclass.
--
-- Parameters:
-- * msg: Email.Message instance
--
-- Returns:
-- * true on success, false on failure
function App:compose(msg)
  self.log.e("compose() not implemented")
  return false
end

--- Email.App:composeFromChooser()
--- Method
--- Given a path, open a chooser and read from the file selected by the user.
--- Use its contents to create an `Email.Message` and then call `compose()`
---
--- Parameters:
--- * `path`: string containing path to directory containing files to offer as options
---
--- Returns:
--- * `Email.Message` instance
function App:composeFromChooser(path)
  local iter, data = hs.fs.dir(path)
  if iter == nil then
    self.log.ef("Cannot read path: %s", path)
    return false
  end
  local choices = {}
  for file in hs.fs.dir(path) do
    -- If filename starts with "." ignore it
    -- This catches "." ".." as well
    if file:sub(1,1) == "." then -- noop
    else
      local choice = {
        ["text"] = file,
        ["path"] = path .. "/" .. file
      }
      table.insert(choices, choice)
    end
  end

  if #choices == 0 then
    self.log.e("No files found in %s", path)
    return false
  end

  table.sort(choices, function(a,b) return a.text:lower() < b.text:lower() end)

  local callback = function(info)
    if not info then
      log.d("User canceled template selection")
    end
    local email = Message.fromFile(info.path)
    if email then
      self:compose(email)
    end
  end

  chooser = hs.chooser.new(callback)
  chooser:choices(choices)
  chooser:show()
end

--- Email.App:useHTMLforCompose()
--- Method
--- Set whether or not to use HTML, as opposed to plain text, when composing.
---
--- Parameters:
--- * `flag`: boolean indicating if HTML should be used
---
--- Returns:
--- * Nothing
function App:useHTMLforCompose(useHTML)
  self.useHTML = flag
end

return App
-- vim: foldmethod=marker: --
