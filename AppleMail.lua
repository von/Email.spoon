--- === Email.AppleMail ===
---
--- An interface for interfacing to Apple Mail
-- Kudos: https://apple.stackexchange.com/questions/125822/applescript-automate-mail-tasks

local AppleMail = {}

-- AppleMail is a subclass of Email.App
local App = dofile(hs.spoons.resourcePath("App.lua"))

-- Failed table lookups on the instances should fallback to the class table, to get methods
AppleMail.__index = AppleMail

setmetatable(AppleMail, {
  -- Failed lookups on class go to superclass
  __index = App
})

AppleMail.AppId = "com.apple.mail"

--- Email.AppleMail.new()
--- Constructor
--- Create a new interface to AppleMail
---
--- Parameters:
--- * None
---
--- Returns:
--- * `Email.AppleMail` instance
function AppleMail.new()
  -- Create a new instance of superclass, but give it metatable of subclass
  local self = setmetatable(App.new(), AppleMail)
  return self
end

--- Email.AppleMail:compose()
--- Method
--- Given an Email.Message instance, create new email composition
--- with its contents.
---
--- Parameters:
--- * `msg`: `Email.Message` instance
---
--- Returns:
--- * `true` on success, `false` on failure
function AppleMail:compose(mail)
  self.log.d("AppleMail:compose() called.")
  local properties = "visible:true, subject:"
  if mail.subject then
    properties = properties .. "\"" .. self:escapeApplescriptString(mail.subject) .. "\""
  end
  if mail.from then
    properties = properties .. ", sender:\"" .. mail.from .. "\""
  end
  if mail.content then
    properties = properties .. ", content:\"" .. self:escapeApplescriptString(mail.content) .. "\""
  end
  local tell_cmds = {}
  if mail.to then
    hs.fnutils.each(mail.to, function(addr) table.insert(tell_cmds, "make new to recipient at newMessage with properties {address:\"" .. self:escapeApplescriptString(addr) .. "\"}") end)
  end
  if mail.cc then
    hs.fnutils.each(mail.cc, function(addr) table.insert(tell_cmds, "make new cc recipient at newMessage with properties {address:\"" .. self:escapeApplescriptString(addr) .. "\"}") end)
  end
  if mail.bcc then
    hs.fnutils.each(mail.bcc, function(addr) table.insert(tell_cmds, "make new bcc recipient at newMessage with properties {address:\"" .. self:escapeApplescriptString(addr) .. "\"}") end)
  end
  if mail.attachment then
    hs.fnutils.each(mail.attachment, function(path) table.insert(tell_cmds, "make new attachment at newMessage with properties {file name:\"" .. self:escapeApplescriptString(path) .. "\"}") end)
  end
  local tell_cmd_str = table.concat(tell_cmds, "\n")
  local script = string.format([[
    tell application "Mail"
      set newMessage to make new outgoing message with properties {%s}
      %s
      activate
    end tell
  ]], properties, tell_cmd_str)
  return self:executeApplescript(script)
end

--- Email.AppleMail:moveToFolder()
--- Method
--- Move current message to folder
--- Mail must be the active application
---
--- Parameters:
--- * folder: string with the name of the target folder, which must appear under
---   Message / Move To
---
--- Returns:
--- * true on success, false on failure
function AppleMail:moveToFolder(folder)
  self.log.df("Moving message to %s", folder)
  local mail = hs.application.find(self.AppId)
  if not mail then
    self.log.e("Could not find Mail application")
    return false
  end
  if not mail:selectMenuItem({"Message", "Move to", folder}) then
    self.log.f("Failed to move message to %s", folder)
    return false
  end
  return true
end

return AppleMail
-- vim: foldmethod=marker: --
