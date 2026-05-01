# DMG PO PDF Agent

On-prem agent that fetches PO PDFs from the local `T:\` drive on demand and uploads them to the `dmgpdfvault` Azure blob container, where the web app's frontend reads them via short-lived SAS URLs. Lets DICI users (and anyone else without T: access) view DMG POs from the browser.

## Architecture (the short version)

```
[browser] --POST--> dmg-followup-proxy --queues request--> popdfrequests table
   |                                                              ^
   | polls /po-pdf-status                                          |
   v                                                               |
[browser]  <----- SAS URL on /po-pdf-status -- proxy generates    |
                                                                   |
                                              this agent --polls--/
                                                  |  /po-agent-poll
                                                  v
                                              walks T:\Année En Cours\...
                                                  |
                                                  v
                                         /po-agent-write-sas
                                              |
                                              v
                                         PUT blob to dmgpdfvault
                                              |
                                              v
                                         /po-agent-complete
```

All traffic from this machine is **outbound HTTPS only**. The agent never accepts inbound connections — it polls the proxy.

## Install (PC11)

1. Open an **elevated PowerShell** (Run as Administrator).
2. `cd` to wherever this folder lives.
3. Run:
   ```powershell
   .\install-agent.ps1
   ```
4. Enter when prompted:
   - Your **DMG portal username + password** (used to authenticate to the proxy)
   - Your **Windows password** (Task Scheduler stores this so the agent can run pre-login)
5. The installer:
   - Tests login against the auth server
   - DPAPI-encrypts portal credentials (CurrentUser scope) → `%ProgramData%\DMG\po-agent\credentials.dat`
   - Copies the agent to `%ProgramData%\DMG\po-agent\`
   - Registers a **Scheduled Task** named `DMG-PO-Agent` that:
     - Runs **AtStartup** as the installing user
     - Uses the saved Windows password to start pre-login
     - Restarts every 1 min on failure

## Uptime / reboot handling

The agent runs **24/7**, even when nobody is logged into Windows. Task Scheduler stores the run-as user's Windows password in the LSA secrets store and uses it to launch PowerShell as that user at boot.

**If you change your Windows password, re-run `install-agent.ps1`** — the saved task password will be stale and the agent will stop launching.

## Why UNC paths instead of `T:\`

Drive-letter mappings (`T:\`) only exist inside an interactive Windows session. A task running pre-login has no `T:` available. The agent uses the equivalent UNC path directly: `\\WIN2K19\BaseDMG\Année En Cours\Bon de Commande`. Same data, no drive mapping required.

## Logs

```
%ProgramData%\DMG\po-agent\logs\agent-YYYY-MM-DD.log
```

Tail the live log:

```powershell
Get-Content "$env:ProgramData\DMG\po-agent\logs\agent-$(Get-Date -Format 'yyyy-MM-dd').log" -Wait -Tail 20
```

## Manual control

```powershell
# Stop / start the task
Stop-ScheduledTask -TaskName 'DMG-PO-Agent'
Start-ScheduledTask -TaskName 'DMG-PO-Agent'

# Check state
Get-ScheduledTask -TaskName 'DMG-PO-Agent' | Get-ScheduledTaskInfo
```

## Uninstall

```powershell
.\uninstall-agent.ps1
```

Removes the task, deletes credentials + agent files. Add `-KeepLogs` to preserve `logs\`.

## Future hardening

- Dedicated `po-agent-pc11` user account in the auth system (instead of using a real user's creds)
- Move from Scheduled Task → proper Windows Service (NSSM-wrapped or rewritten in C#) for stricter lifecycle
- IP-pin the agent's bearer token to PC11's outbound IP
- Periodic `auth-login` rotation
