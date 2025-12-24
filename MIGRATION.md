# MCReporting Migration Summary

## Replacements Performed

| File | Line | Old Value | New Value |
|------|------|-----------|-----------|
| MfEmail.asp | 84 | `192.168.20.32` | `192.168.20.3` |

## Values Found But NOT In Replacement Table

These IPs/values were found but are NOT in the migration table - no changes made:

| Value | Files | Context |
|-------|-------|---------|
| 192.168.20.50 | SystemError.asp:40, MfEmail.asp:122,224, HsEmail.asp:94,157, NcEmail.asp:94,157, HsSystemError.asp:40 | SMTP session conditional checks |
| 192.168.20.37 | SystemError.asp:40, MfEmail.asp:122,224, HsEmail.asp:94,157, NcEmail.asp:94,157 | SMTP session conditional checks |
| 192.168.20.110 | UsersDormant.asp:36, UsersOnline.asp:36 | Developer PC IP for testing |
| 192.168.20.85 | JsFiles/ClarityJSFunc.js:266 | Redo web location URL |
| 192.168.20.1 | JsFiles/ClarityJSFunc.js:258,261,279 | Commented out code |
| 192.168.20.205 | JsFiles/ClarityJSFunc.js:264,281,303,304 | Commented out code |

## Unchanged Values (Per Migration Table)

| Value | Files | Status |
|-------|-------|--------|
| 82.71.163.186 | MfEmail.asp:86 | External IP - unchanged per spec |
| mx496502.smtp-engine.com | MfEmail.asp, HsEmail.asp, NcEmail.asp, IhrEmail.asp, SendEmailReminder.asp | Email host - unchanged per spec |

## Registry Keys

No registry key references found in this ASP Classic web application.

## DLL Paths

No DLL references found in this ASP Classic web application.

## Dependencies

- JMail.Message COM object (email sending)
- ADODB.Recordset (database access)
- Scripting.FileSystemObject (file operations)
- ADODB.Stream (binary file handling)

## Database Connections

Connection strings stored in Application/Session variables:
- `Application("ConnMcLogon")`
- `Session("ConnMachinefaults")`

Actual connection strings likely configured externally (IIS or include files).

## Notes

- Line 86 in MfEmail.asp has `http://http://` typo (not fixed - outside migration scope)
- Database paths using `192.168.20.36` were NOT found in this repository
- Server names (mc-fp, mc-sqlsvr, mc-test, mc-host) were NOT found in this repository
