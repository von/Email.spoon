--- === Email.Outlook ===
---
--- Class for interfacing to Microsoft Outlook
-- Kudos: https://apple.stackexchange.com/questions/125822/applescript-automate-mail-tasks

local Outlook = {}

-- AppleMail is a subclass of Email.App
local App = dofile(hs.spoons.resourcePath("App.lua"))

-- Failed table lookups on the instances should fallback to the class table, to get methods
Outlook.__index = Outlook

setmetatable(Outlook, {
  -- Failed lookups on class go to superclass
  __index = App
})

Outlook.AppId = "com.microsoft.Outlook"

function Outlook.new()
  -- Create a new instance of superclass, but give it metatable of subclass
  local self = setmetatable(App.new(), Outlook)
  return self
end

--- Email.Outlook:compose()
--- Method
--- Given an Email.Message instance, create new email comosition
--- with its contents.
---
--- Parameters:
--- * msg: Email.Message instance
---
--- Returns:
--- * true on success, false on failure
function Outlook:compose(mail)
  self.log.d("Outlook:compose() called.")
  local properties = "subject:"
  if mail.subject then
    properties = properties .. "\"" .. mail.subject .. "\""
  end
  if mail.from then
    properties = properties .. ", sender:\"" .. mail.from .. "\""
  end
  if mail.content then
    properties = properties .. ", plain text content:\"" .. self:scapeApplescriptString(mail.content) .. "\""
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
  local tell_cmd_str = ""
  hn.fnutils(tell_cmds, function(c) tell_cmd_str = tell_cmd_str .. c .. "\n" end)
  local script = string.format([[
    tell application "Outlook"
      set theMessage to make new outgoing message with properties {%s}
      %s
      open theMessage
      activate
    end tell
  ]], properties, tell_cmd_str)
  return self:executeApplescript(script)
end
--- Method
--- Move current message to archive
--- Uses Outlook's existing ^E shortcut
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:moveToArchive()
  self.log.d("Moving message to archive")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  hs.eventtap.event.newKeyEvent({"ctrl"}, "e", true):post(outlook)
end
-- }}} moveToArchive --

-- moveToFolder {{{ --
--- Email.Outlook:moveToFolder()
--- Method
--- Move current message to folder
--- Mail must be the active application
---
--- Parameters:
--- * `folder`: string with the name of the target folder, which must appear under
---   `Message / Move To`
---
--- Returns:
--- * Nothing
function Outlook:moveToFolder(folder)
  -- Kudos: https://apple.stackexchange.com/a/213044/104604
  self.log.df("Moving message to %s", folder)
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Move", folder}) then
    -- Some versions of Outlook use "<foldername> (<email address>)"
    if not outlook:selectMenuItem(folder .. " \\(.*\\)", true) then
      self.log.f("Failed to move message to %s", folder)
    end
  end
end
-- }}} moveToFolder --

-- flag() {{{ --
--- Email.Outlook:flag()
--- Method
--- Flag message for followup
--- Email.Outlook must be the active application
---
--- Parameters:
--- * `when` (optional): a string and one of: `Today`, `Tomorrow`, `This Week`,
---   `Next Week`, `No Due Date`. Default is `Today`
---
--- Returns:
--- * Nothing
function Outlook:flag(when)
  when = when or "Today"
  self.log.df("Flagging message for followup %s", when)
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Follow Up", when}) then
    self.log.f("Failed to flag message for %s", when)
  end
end
-- }}} flag() --

-- clearFlag() {{{ --
--- Email.Outlook:clearFlag()
--- Method
--- Clear any flag on message for followup
--- Email.Outlook must be the active application
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:clearFlag(when)
  self.log.d("Clearing flagging on message for followup")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Follow Up", "Clear Flag"}) then
    self.log.i("Failed to clear flag on message")
  end
end
-- }}} clearFlag() --

-- reply() {{{ --
--- Email.Outlook:reply()
--- Method
--- Reply to current message
--- Email.Outlook must be the active application
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:reply()
  log.d("Replying to message")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Reply"}) then
    self.log.f("Failed to reply to message")
  end
end
-- }}} reply() --

-- replyAll() {{{ --
--- Email.Outlook:replyAll()
--- Method
--- Reply-all to current message
--- Email.Outlook must be the active application
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:replyAll()
  self.log.d("Replying all to message")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Reply All"}) then
    self.log.f("Failed to reply all to message")
  end
end
-- }}} replyAll() --

-- forward() {{{ --
--- Email.Outlook:forward()
--- Method
--- Forward current message
--- Email.Outlook must be the active application
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:forward()
  self.log.d("Forwarding message")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Forward"}) then
    self.log.f("Failed to forward message")
  end
end
-- }}} forward() --

return Outlook
-- vim: foldmethod=marker: --
