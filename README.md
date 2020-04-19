# o365mailrules

Using Microsoft Graph API to pull all users mail rules.

FYI: I'm not a developer by trade :) so feedback is always welcome if you see ugly or inefficient code.

Note this only pulls all mail rules (which can be a forward rule). Not the global mail forwarding setting. as of 4/19/2020 There doesn't appear to be a way to pull those via the MS Graph API. So you have to use powershell still. Pulling Global foward setting via powershell is fast where Mail rules takes forever. This script takes about 7 minutes for just over 37,000 users. Granted many of those don't have mail rules (but that actually slows the API call down)

Created with Python 3.7.4. Definitely  won't work with 3.6. Not tested below 3.7.4.
Coded up on Windows and runs on Linux for my reasons.

Requires "requests" and "cryptography"

Don't forget to change the folder locations in conf.ini to handle windows or linux.

I'm fully aware I only obfuscated the client secret stored in the "keyFile". I just didn't want it to be in plaintext and I wanted to mess around with the "cryptography" library. After the Python Scripts first run it will encode the passwords under the [keys] section based on the "fileKey" in the conf.ini and a salt in the script. 

I forget if you need to pre create the "outputDir" and the "LogDir"

There is NO builtin cleanup out outputted files.

The script is hard coded set to run 36 threads and "20" queries per a batch API call to Microsoft Graph. That is the max queries that can be added to a single batch. I'll get around to making it a setting in the conf.ini. With just over 37,000 users it didn't not hit the Graph API limits of 10,000 queries per 10 minutes. I belive that is the limit for Outlook queries. this is doing Mail Rules so i'm guessing that would apply.

very little exception handling. That is always the last to thing be added isn't it?

Don't forget to register an App within Azure AD and give it Microsoft Graph API rights. Should just need Directory.Read.All, MailboxSettings.Read, User.Read.All.

Also uses Client Secret only to connect. No Certificate. 

That's all I can think of right now. Any questions just ask me.

