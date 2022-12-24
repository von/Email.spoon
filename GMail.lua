--- === Email.GMail ===
---
--- An interface for GMail

local GMail = {}

-- GMail is a subclass of Email.App
local App = dofile(hs.spoons.resourcePath("App.lua"))

-- Failed table lookups on the instances should fallback to the class table, to get methods
GMail.__index = GMail

setmetatable(GMail, {
  -- Failed lookups on class go to superclass
  __index = App
})

-- {{{ urlencode()
-- Kudos: https://gist.github.com/liukun/f9ce7d6d14fa45fe9b924a3eed5c3d99
local char_to_hex = function(c)
  return string.format("%%%02X", string.byte(c))
end

local function urlencode(url)
  if url == nil then
    return
  end
  url = url:gsub("\n", "\r\n")
  url = url:gsub("([^%w ])", char_to_hex)
  url = url:gsub(" ", "+")
  return url
end
-- }}} urlencode()

--- Email.AppleMail.new()
--- Constructor
--- Create a new interface to AppleMail
---
--- Parameters:
--- * None
---
--- Returns:
--- * `Email.GMail` instance
function GMail.new()
  -- Create a new instance of superclass, but give it metatable of subclass
  local self = setmetatable(App.new(), GMail)
  return self
end

--- Email.GMail:compose()
--- Method
--- Given an Email.Message instance, create new email composition
--- with its contents.
---
--- Parameters:
--- * `msg`: `Email.Message` instance
---
--- Returns:
--- * `true` on success, `false` on failure
function GMail:compose(mail)
  self.log.d("GMail:compose() called.")
  -- Create URL. Kudos: https://stackoverflow.com/a/8852679/197789
  local url = "https://mail.google.com/mail/?view=cm&fs=1"
  if mail.to then
    url = url .. "&to=" .. mail.to
  end
  if mail.cc then
    url = url .. "&cc=" .. mail.cc
  end
  if mail.bcc then
    url = url .. "&bcc=" .. mail.bcc
  end
  -- TODO: handle mail.from
  if mail.subject then
    url = url .. "&su=" .. urlencode(mail.subject)
  end
  if mail.content then
    url = url .. "&body=" .. urlencode(mail.content)
  end
  -- TODO: handle mail.attachment

  return self:open(url)
end


--- Email.GMail:open()
--- Method
--- Given a url, open it. Uses hs.urlevent.openURL(url).
---
--- Parameters:
--- * `url`: URL to open
---
--- Returns:
--- * `true` on success, `false` on failure
function GMail:open(url)
  self.log.df("GMail:open(%s) called.", url)
  return hs.urlevent.openURL(url)
end

return GMail
-- vim: foldmethod=marker: --
