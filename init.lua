--- === Email ===
---
--- An interface to creating and sending emails and email applications.
local Email = {}

-- Metadata {{{ --
Email.name="Email"
Email.version="0.1"
Email.author="Von Welch"
-- https://opensource.org/licenses/Apache-2.0
Email.license="Apache-2.0"
Email.homepage="https://github.com/von/Email.spoon"
-- }}} Metadata --

-- Failed table lookups on the instances should fallback to the class table, to get methods
Email.__index = Email

Email.log = hs.logger.new("Email")

--- Email:debug()
--- Method
--- Enable or disable debugging
---
--- Parameters:
--- * enable: a boolean to indiciate whether debigging should be enabled or disabled
---
--- Returns:
--- * Nothing
function Email:debug(enable)
  if enable then
    self.log.setLogLevel('debug')
    self.log.d("Debugging enabled")
  else
    self.log.d("Disabling debugging")
    self.log.setLogLevel('info')
  end
end

Email.Message = dofile(hs.spoons.resourcePath("Message.lua"))

--- Email.AppleMail()
--- Constructor
--- Create an interface to Apple Mail.
---
--- Parameters:
--- * None
---
--- Returns:
--- * Email.AppleMail instance
function Email.AppleMail()
  return dofile(hs.spoons.resourcePath("AppleMail.lua")).new()
end

--- Email.GMail()
--- Constructor
--- Create an interface to GMail.
---
--- Parameters:
--- * None
---
--- Returns:
--- * Email.GMail instance
function Email.GMail()
  return dofile(hs.spoons.resourcePath("GMail.lua")).new()
end

--- Email.Outlook()
--- Constructor
--- Create an interface to Outlook.
---
--- Parameters:
--- * None
---
--- Returns:
--- * Email.Outlook instance
function Email.Outlook()
  return dofile(hs.spoons.resourcePath("Outlook.lua")).new()
end

return Email
-- vim: foldmethod=marker: --
