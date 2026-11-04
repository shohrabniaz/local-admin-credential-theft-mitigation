# Local administrator rights and credential-theft mitigation for SMEs

Fictional case: Coastal Managed IT-style Windows endpoints where local admin rights and shared local passwords are common. **Defensive laboratory only.**

## Research questions

1. What local-privilege and credential-theft exposure exists on a typical SME endpoint?
2. Which mitigations most effectively reduce lateral-movement risk?
3. What is the impact on users and IT support?

## Deliverables

- Endpoint privilege / credential exposure assessment (live inventory + scored model)
- Mitigation design (least privilege, Windows LAPS, Credential Guard / ASR where available)
- Lab validation results (control coverage — no attack-path tooling)
- SME hardening guide, report, slides

## Scripts

```powershell
.\01_Inventory_Local_Admins.ps1
.\02_Check_Laps_And_Credential_Guard.ps1
.\03_Score_Exposure_And_Mitigations.ps1
```

## Ethics

No credential dumping, no BloodHound-style attack-path tooling, no lateral-movement procedures. Exposure is measured by configuration inventory and Essential Eight–aligned control coverage.
