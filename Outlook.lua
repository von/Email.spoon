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

-- new() {{{ --
--- Email.Outlook.new()
--- Constructor
--- Create a new instance of an Outlook interface.
---
--- Parameters:
--- * None
---
--- Returns:
--- * Email.Outlook instance
function Outlook.new()
  -- Create a new instance of superclass, but give it metatable of subclass
  local self = setmetatable(App.new(), Outlook)
  return self
end
-- }}} Outlook.new() --

-- compose() {{{ --
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
    properties = properties .. "\"" .. self:escapeApplescriptString(mail.subject) .. "\""
  end
  if mail.from then
    properties = properties .. ", sender:\"" .. mail.from .. "\""
  end
  if mail.content then
    local content = self:escapeApplescriptString(mail.content)
    if self.useHTML then
      -- "New" Outlook expects content in HTML, so add a <br> whereever we have a CR
      -- Kudos: https://discussions.apple.com/thread/5929457
      content = content:gsub("\n", "<br>%1")
    end
    properties = properties .. ", plain text content:\"" .. content .. "\""
  end
  local emailToStr = function(addr) return string.format("{email address:{address:\"%s\"}}", self:escapeApplescriptString(addr)) end
  local tell_cmds = {}
  if mail.to then
    hs.fnutils.each(mail.to, function(addr) table.insert(tell_cmds, "make new to recipient at newMessage with properties " .. emailToStr(addr)) end)
  end
  if mail.cc then
    hs.fnutils.each(mail.cc, function(addr) table.insert(tell_cmds, "make new cc recipient at newMessage with properties " .. emailToStr(addr)) end)
  end
  if mail.bcc then
    hs.fnutils.each(mail.bcc, function(addr) table.insert(tell_cmds, "make new bcc recipient at newMessage with properties " .. emailToStr(addr)) end)
  end
  if mail.attachment then
    hs.fnutils.each(mail.attachment, function(path) table.insert(tell_cmds, "make new attachment at newMessage with properties {file name:\"" .. self:escapeApplescriptString(path) .. "\"}") end)
  end
  local tell_cmd_str = table.concat(tell_cmds, "\n")
  local script = string.format([[
    tell application "Outlook"
      set newMessage to make new outgoing message with properties {%s}
      %s
      open newMessage
      activate
    end tell
  ]], properties, tell_cmd_str)
  return self:executeApplescript(script)
end
-- }}} Email.Outlook:compose() --

-- {{{ moveToArchive() --
--- Email.Outlook:moveToArchive()
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

-- moveToFolder() {{{ --
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
      hs.alert("Failed to move message to " .. folder)
    end
  end
end
-- }}} moveToFolder --

-- flag() {{{ --
--- Email.Outlook:flag()
--- Method
--- Flag current message
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:flag()
  self.log.d("Flagging message")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Flag"}) then
    self.log.e("Failed to flag to message")
  end
end
-- }}} flag() --

-- clearFlag() {{{ --
--- Email.Outlook:clearFlag()
--- Method
--- Unflag message.
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:clearFlag(when)
  self.log.d("Unflagging message")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Message", "Unflag"}) then
    self.log.e("Failed to unflag to message")
  end
end
-- }}} clearFlag() --

-- snooze() {{{ --
--- Email.Outlook:snooze()
--- Method
--- Snooze current message.
--- XXX This doesn't work right now because I cannot find the menu item.
---
--- Parameters:
--- * when (optional): Until when to snooze message. Parameter is a string and must
---   match menu item under "Messages / Snooze". Default is "Tomorrow"
---
--- Returns:
--- * Nothing
function Outlook:snooze(when)
  when = when or "Tomorrow"
  self.log.df("Snooze message until %s", when)
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  -- The Outlook menu for Snoomze (under Message / Snooze) appends the actual date/time
  -- to the base string, so we need to use a regex.
  -- XXX I cannot get hammerspoon to find the Snooze menu items...
  if not outlook:selectMenuItem(when .. " .*", true) then
    self.log.ef("Failed to snooze message to %s", when)
  end
end
-- }}} snooze() --

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

-- delete() {{{ --
--- Email.Outlook:delete()
--- Method
--- Delete current message
---
--- Parameters:
--- * None
---
--- Returns:
--- * Nothing
function Outlook:delete()
  self.log.d("Deleting message")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return
  end
  if not outlook:selectMenuItem({"Edit", "Delete"}) then
    self.log.f("Failed to delete message")
  end
end
-- }}} delete() --

-- focusOnCalendar() {{{ --
--- Email.Outlook:focusOnCalendar()
--- Method
--- Focus on the calendar window.
---
--- Parameters:
--- * None
---
--- Returns:
--- * True on success, false on failure
function Outlook:focusOnCalendar()
  self.log.d("Focusing on Calendar")
  local outlook = hs.application.find(self.AppId)
  if not outlook then
    self.log.e("Could not find Outlook application")
    return false
  end
  local win = outlook:findWindow("Calendar")
  if not win then
    self.log.e("Could not find Calendar window")
    return false
  end
  win:focus()
  return true
end
-- }}} focusOnCalendar()

return Outlook
-- vim: foldmethod=marker: --
