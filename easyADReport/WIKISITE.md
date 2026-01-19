# 🛡️ ACTIVE DIRECTORY SECURITY & KOMPROMITTIERUNG - MICROSOFT LEARN WIKI

## 📚 INHALTSVERZEICHNIS

1. [Kerberos-Angriffe](#kerberos-angriffe)
2. [Credential Theft](#credential-theft)
3. [Privilege Escalation](#privilege-escalation)
4. [Lateral Movement](#lateral-movement)
5. [Persistence](#persistence)
6. [Certificate Services Angriffe](#certificate-services-angriffe)
7. [GPO-Angriffe](#gpo-angriffe)
8. [LDAP & Directory Services](#ldap--directory-services)
9. [Monitoring & Detection](#monitoring--detection)
10. [Hardening & Best Practices](#hardening--best-practices)

---

## 🎫 KERBEROS-ANGRIFFE

### Kerberoasting
**Beschreibung:** Angriff auf Service Principal Names (SPNs) um Service-Account-Passwörter offline zu cracken

**Microsoft Learn Ressourcen:**
- [Kerberos Authentication Overview](https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-authentication-overview)
- [Service Principal Names (SPNs)](https://learn.microsoft.com/en-us/windows/win32/ad/service-principal-names)
- [Kerberos Constrained Delegation](https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-constrained-delegation-overview)

**Mitigations:**
- [Managed Service Accounts](https://learn.microsoft.com/en-us/windows-server/security/group-managed-service-accounts/group-managed-service-accounts-overview)
- [Strong Password Policies](https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/password-policy)

**Detection:**
- Event ID 4769 (Kerberos Service Ticket Request)
- [Advanced Audit Policy Configuration](https://learn.microsoft.com/en-us/windows/security/threat-protection/auditing/advanced-security-audit-policy-settings)

---

### AS-REP Roasting
**Beschreibung:** Angriff auf Accounts ohne Kerberos Pre-Authentication

**Microsoft Learn Ressourcen:**
- [Kerberos Pre-Authentication](https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/network-security-configure-encryption-types-allowed-for-kerberos)
- [User Account Control Flags](https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/useraccountcontrol-manipulate-account-properties)

**Mitigations:**
- [Disable "Do not require Kerberos preauthentication"](https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/accounts-require-kerberos-preauthentication)

**Detection:**
- Event ID 4768 (Kerberos TGT Request)

---

### Golden Ticket
**Beschreibung:** Forged Kerberos TGT mit KRBTGT-Hash

**Microsoft Learn Ressourcen:**
- [KRBTGT Account](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/ad-forest-recovery-resetting-the-krbtgt-password)
- [Kerberos Ticket Lifetime](https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/maximum-lifetime-for-service-ticket)

**Mitigations:**
- [Reset KRBTGT Password](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/ad-forest-recovery-resetting-the-krbtgt-password)
- [Protected Users Security Group](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/protected-users-security-group)

**Detection:**
- Event ID 4768, 4769 (Unusual TGT/Service Ticket patterns)
- [Microsoft Defender for Identity](https://learn.microsoft.com/en-us/defender-for-identity/what-is)

---

### Silver Ticket
**Beschreibung:** Forged Kerberos Service Ticket

**Microsoft Learn Ressourcen:**
- [Kerberos Service Tickets](https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-authentication-overview)
- [Service Account Security](https://learn.microsoft.com/en-us/windows-server/security/group-managed-service-accounts/getting-started-with-group-managed-service-accounts)

**Mitigations:**
- [Managed Service Accounts](https://learn.microsoft.com/en-us/windows-server/security/group-managed-service-accounts/group-managed-service-accounts-overview)
- Regular Service Account Password Rotation

---

### Unconstrained Delegation
**Beschreibung:** Computer/User mit Unconstrained Delegation können TGTs von Benutzern speichern

**Microsoft Learn Ressourcen:**
- [Kerberos Delegation](https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-constrained-delegation-overview)
- [Resource-Based Constrained Delegation](https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-constrained-delegation-overview#resource-based-constrained-delegation)

**Mitigations:**
- [Account is sensitive and cannot be delegated](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/component-updates/tgt-delegation)
- Use Constrained Delegation instead

**Detection:**
- Event ID 4624 (Logon with delegation)

---

## 🔑 CREDENTIAL THEFT

### Pass-the-Hash (PtH)
**Beschreibung:** Verwendung von NTLM-Hashes statt Klartext-Passwörtern

**Microsoft Learn Ressourcen:**
- [NTLM Overview](https://learn.microsoft.com/en-us/windows-server/security/kerberos/ntlm-overview)
- [Mitigating Pass-the-Hash Attacks](https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/network-security-restrict-ntlm-ntlm-authentication-in-this-domain)

**Mitigations:**
- [Credential Guard](https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/credential-guard)
- [Protected Users Group](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/protected-users-security-group)
- [Disable NTLM](https://learn.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/network-security-restrict-ntlm-ntlm-authentication-in-this-domain)

**Detection:**
- Event ID 4624 (Logon Type 3 with NTLM)
- [Advanced Threat Analytics](https://learn.microsoft.com/en-us/advanced-threat-analytics/what-is-ata)

---

### Pass-the-Ticket (PtT)
**Beschreibung:** Verwendung gestohlener Kerberos-Tickets

**Microsoft Learn Ressourcen:**
- [Kerberos Ticket Cache](https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-authentication-overview)

**Mitigations:**
- [Credential Guard](https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/credential-guard)
- [Remote Credential Guard](https://learn.microsoft.com/en-us/windows/security/identity-protection/remote-credential-guard)

---

### DCSync
**Beschreibung:** Missbrauch von Directory Replication Permissions

**Microsoft Learn Ressourcen:**
- [Active Directory Replication](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/get-started/replication/active-directory-replication-concepts)
- [Directory Service Access Auditing](https://learn.microsoft.com/en-us/windows/security/threat-protection/auditing/audit-directory-service-access)

**Mitigations:**
- Restrict Replication Permissions
- [AdminSDHolder Protection](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory)

**Detection:**
- Event ID 4662 (Directory Service Access)
- [Microsoft Defender for Identity](https://learn.microsoft.com/en-us/defender-for-identity/what-is)

---

### LSASS Memory Dumping
**Beschreibung:** Credential Extraction aus LSASS-Prozess

**Microsoft Learn Ressourcores:**
- [LSA Protection](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection)
- [Credential Guard](https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/credential-guard)

**Mitigations:**
- [RunAsPPL for LSASS](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/configuring-additional-lsa-protection)
- [Credential Guard](https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/credential-guard-manage)

**Detection:**
- Event ID 4656 (Handle to LSASS)
- [Windows Defender Application Control](https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-defender-application-control/windows-defender-application-control)

---

## ⬆️ PRIVILEGE ESCALATION

### AdminSDHolder Abuse
**Beschreibung:** Persistenz durch Manipulation von AdminSDHolder

**Microsoft Learn Ressourcen:**
- [AdminSDHolder and SDProp](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory)
- [Protected Accounts and Groups](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory)

**Mitigations:**
- Regular AdminSDHolder Auditing
- [Audit Directory Service Changes](https://learn.microsoft.com/en-us/windows/security/threat-protection/auditing/audit-directory-service-changes)

**Detection:**
- Event ID 5136 (Directory Service Object Modified)

---

### GPO Modification
**Beschreibung:** Privilege Escalation durch GPO-Manipulation

**Microsoft Learn Ressourcen:**
- [Group Policy Security](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory)
- [Delegating Group Policy Management](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-r2-and-2008/cc731076(v=ws.10))

**Mitigations:**
- Restrict GPO Modification Rights
- [GPO Auditing](https://learn.microsoft.com/en-us/windows/security/threat-protection/auditing/audit-policy-change)

**Detection:**
- Event ID 5136, 5137, 5141 (GPO Changes)

---

### DCShadow
**Beschreibung:** Rogue Domain Controller für AD-Manipulation

**Microsoft Learn Ressourcen:**
- [Domain Controller Security](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/securing-domain-controllers-against-attack)

**Mitigations:**
- Monitor Domain Controller Registrations
- [Network Segmentation](https://learn.microsoft.com/en-us/security/zero-trust/deploy/networks)

**Detection:**
- Event ID 4742 (Computer Account Changed)
- [Microsoft Defender for Identity](https://learn.microsoft.com/en-us/defender-for-identity/dcshadow-attack)

---

### ACL Abuse
**Beschreibung:** Missbrauch von Permissions (GenericAll, WriteDACL, etc.)

**Microsoft Learn Ressourcen:**
- [Access Control Lists](https://learn.microsoft.com/en-us/windows/win32/secauthz/access-control-lists)
- [Active Directory Permissions](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-d--securing-built-in-administrator-accounts-in-active-directory)

**Mitigations:**
- Regular Permission Audits
- [Least Privilege Principle](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/implementing-least-privilege-administrative-models)

**Detection:**
- Event ID 5136 (Directory Service Object Modified)

---

## 🔄 LATERAL MOVEMENT

### WMI/DCOM Abuse
**Beschreibung:** Remote Code Execution via WMI/DCOM

**Microsoft Learn Ressourcen:**
- [WMI Security](https://learn.microsoft.com/en-us/windows/win32/wmisdk/securing-wmi-namespaces)
- [DCOM Security](https://learn.microsoft.com/en-us/windows/win32/com/dcom-security-enhancements)

**Mitigations:**
- [Windows Firewall](https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-firewall/windows-firewall-with-advanced-security)
- Disable WMI/DCOM where not needed

**Detection:**
- Event ID 4648 (Explicit Credentials)

---

### PsExec/Remote Services
**Beschreibung:** Remote Command Execution

**Microsoft Learn Ressourcen:**
- [Service Control Manager](https://learn.microsoft.com/en-us/windows/win32/services/service-control-manager)

**Mitigations:**
- [Local Administrator Password Solution (LAPS)](https://learn.microsoft.com/en-us/windows-server/identity/laps/laps-overview)
- Restrict Service Creation Rights

**Detection:**
- Event ID 7045 (Service Installation)
- Event ID 4697 (Service Installed)

---

### RDP Hijacking
**Beschreibung:** Session Hijacking via RDP

**Microsoft Learn Ressourcen:**
- [Remote Desktop Services Security](https://learn.microsoft.com/en-us/windows-server/remote/remote-desktop-services/rds-security)

**Mitigations:**
- [Network Level Authentication](https://learn.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/remote-desktop-allow-access)
- [Restricted Admin Mode](https://learn.microsoft.com/en-us/windows/security/identity-protection/remote-credential-guard)

**Detection:**
- Event ID 4778, 4779 (Session Reconnect/Disconnect)

---

## 🔐 PERSISTENCE

### Skeleton Key
**Beschreibung:** Backdoor in Domain Controllers

**Microsoft Learn Ressourcen:**
- [Domain Controller Security](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/securing-domain-controllers-against-attack)

**Mitigations:**
- [Credential Guard on DCs](https://learn.microsoft.com/en-us/windows/security/identity-protection/credential-guard/credential-guard)
- Regular DC Integrity Checks

**Detection:**
- Event ID 4673 (Sensitive Privilege Use)
- Memory Analysis

---

### SID History Injection
**Beschreibung:** Privilege Escalation via SID History

**Microsoft Learn Ressourcen:**
- [SID History](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/component-updates/sid-history)
- [SID Filtering](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/component-updates/sid-filtering)

**Mitigations:**
- Enable SID Filtering
- Regular SID History Audits

**Detection:**
- Event ID 4765 (SID History Added)

---

### Directory Service Restore Mode (DSRM) Abuse
**Beschreibung:** Backdoor via DSRM Account

**Microsoft Learn Ressourcen:**
- [Directory Services Restore Mode](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/ad-forest-recovery-resetting-the-krbtgt-password)

**Mitigations:**
- Strong DSRM Passwords
- Regular Password Changes
- Disable Network Logon for DSRM

**Detection:**
- Event ID 4794 (DSRM Password Change Attempt)

---

## 📜 CERTIFICATE SERVICES ANGRIFFE

### ESC1 - Misconfigured Certificate Templates
**Beschreibung:** Certificate Templates mit gefährlichen Permissions

**Microsoft Learn Ressourcen:**
- [Active Directory Certificate Services](https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview)
- [Certificate Template Security](https://learn.microsoft.com/en-us/windows-server/networking/core-network-guide/cncg/server-certs/configure-server-certificate-autoenrollment)

**Mitigations:**
- Review Certificate Template Permissions
- Disable Auto-Enrollment where not needed
- [Certificate Template Hardening](https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/certificate-template-security)

---

### ESC2-ESC8 (Certified Pre-Owned)
**Beschreibung:** Verschiedene AD CS Schwachstellen

**Microsoft Learn Ressourcen:**
- [AD CS Security Best Practices](https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview)
- [Certificate Enrollment Web Services](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/hh831822(v=ws.11))

**Mitigations:**
- Regular AD CS Security Audits
- Restrict Certificate Enrollment Permissions

---

## 🔒 GPO-ANGRIFFE

### GPO Hijacking
**Beschreibung:** Übernahme von GPO-Kontrolle

**Microsoft Learn Ressourcen:**
- [Group Policy Security](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory)

**Mitigations:**
- Restrict GPO Modification Rights
- [GPO Delegation Best Practices](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-r2-and-2008/cc731076(v=ws.10))

---

### Startup/Shutdown Script Abuse
**Beschreibung:** Malicious Scripts via GPO

**Microsoft Learn Ressourcen:**
- [Group Policy Scripts](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/dn789196(v=ws.11))

**Mitigations:**
- [AppLocker](https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-defender-application-control/applocker/applocker-overview)
- Script Signing Requirements

---

## 📂 LDAP & DIRECTORY SERVICES

### LDAP Relay Attacks
**Beschreibung:** Relay-Angriffe auf LDAP

**Microsoft Learn Ressourcen:**
- [LDAP Channel Binding](https://learn.microsoft.com/en-us/windows/win32/ad/ldap-channel-binding)
- [LDAP Signing](https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/enable-ldap-signing-in-windows-server)

**Mitigations:**
- [Enable LDAP Signing](https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/enable-ldap-signing-in-windows-server)
- [LDAP Channel Binding](https://learn.microsoft.com/en-us/windows/win32/ad/ldap-channel-binding)

---

### Anonymous LDAP Binds
**Beschreibung:** Unauthenticated LDAP Access

**Microsoft Learn Ressourcen:**
- [LDAP Security](https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2003/cc773257(v=ws.10))

**Mitigations:**
- Disable Anonymous LDAP Binds
- [LDAP over SSL](https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/enable-ldap-over-ssl-with-3rd-certification-authority)

---

## 🔍 MONITORING & DETECTION

### Windows Event Logs
**Microsoft Learn Ressourcen:**
- [Security Auditing](https://learn.microsoft.com/en-us/windows/security/threat-protection/auditing/security-auditing-overview)
- [Advanced Audit Policy](https://learn.microsoft.com/en-us/windows/security/threat-protection/auditing/advanced-security-audit-policy-settings)

**Wichtige Event IDs:**
- 4624 - Account Logon
- 4625 - Failed Logon
- 4648 - Explicit Credentials
- 4662 - Directory Service Access
- 4672 - Special Privileges Assigned
- 4768 - Kerberos TGT Request
- 4769 - Kerberos Service Ticket
- 4776 - NTLM Authentication
- 5136 - Directory Service Object Modified
- 7045 - Service Installation

---

### Microsoft Defender for Identity
**Microsoft Learn Ressourcen:**
- [Defender for Identity Overview](https://learn.microsoft.com/en-us/defender-for-identity/what-is)
- [Defender for Identity Alerts](https://learn.microsoft.com/en-us/defender-for-identity/alerts-overview)
- [Defender for Identity Architecture](https://learn.microsoft.com/en-us/defender-for-identity/architecture)

---

### Microsoft Sentinel
**Microsoft Learn Ressourcen:**
- [Microsoft Sentinel Overview](https://learn.microsoft.com/en-us/azure/sentinel/overview)
- [Active Directory Data Connector](https://learn.microsoft.com/en-us/azure/sentinel/data-connectors-reference)

---

## 🛡️ HARDENING & BEST PRACTICES

### Tiered Administration Model
**Microsoft Learn Ressourcen:**
- [Enterprise Access Model](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-access-model)
- [Tier 0 Assets](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-security-levels)

---

### Privileged Access Workstations (PAW)
**Microsoft Learn Ressourcen:**
- [Privileged Access Workstations](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-devices)
- [PAW Deployment](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-deployment)

---

### Just-In-Time (JIT) Administration
**Microsoft Learn Ressourcen:**
- [Privileged Identity Management](https://learn.microsoft.com/en-us/azure/active-directory/privileged-identity-management/pim-configure)
- [JIT Access](https://learn.microsoft.com/en-us/azure/defender-for-cloud/just-in-time-access-usage)

---

### Local Administrator Password Solution (LAPS)
**Microsoft Learn Ressourcen:**
- [Windows LAPS Overview](https://learn.microsoft.com/en-us/windows-server/identity/laps/laps-overview)
- [Windows LAPS Deployment](https://learn.microsoft.com/en-us/windows-server/identity/laps/laps-scenarios-windows-server-active-directory)

---

### Protected Users Security Group
**Microsoft Learn Ressourcen:**
- [Protected Users Group](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/protected-users-security-group)
- [How Protected Users Group Works](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/how-credentials-protection-works)

---

### Authentication Policies and Silos
**Microsoft Learn Ressourcen:**
- [Authentication Policies](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/authentication-policies-and-authentication-policy-silos)
- [Configure Authentication Policies](https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/how-to-configure-protected-accounts)

---

## 📖 ZUSÄTZLICHE RESSOURCEN

### Microsoft Security Documentation
- [Windows Security Baselines](https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-security-configuration-framework/windows-security-baselines)
- [Security Compliance Toolkit](https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-security-configuration-framework/security-compliance-toolkit-10)
- [CIS Benchmarks Integration](https://learn.microsoft.com/en-us/compliance/regulatory/offering-cis-benchmark)

### Attack Frameworks
- [MITRE ATT&CK for Enterprise](https://attack.mitre.org/matrices/enterprise/)
- [Microsoft Cybersecurity Reference Architectures](https://learn.microsoft.com/en-us/security/cybersecurity-reference-architecture/mcra)

### Training & Certifications
- [Microsoft Security Training](https://learn.microsoft.com/en-us/training/browse/?terms=security)
- [SC-200: Microsoft Security Operations Analyst](https://learn.microsoft.com/en-us/certifications/exams/sc-200)
- [SC-300: Microsoft Identity and Access Administrator](https://learn.microsoft.com/en-us/certifications/exams/sc-300)

---

## 🔗 QUICK REFERENCE LINKS

### Security Tools
- [PingCastle](https://www.pingcastle.com/) - AD Security Assessment
- [BloodHound](https://github.com/BloodHoundAD/BloodHound) - AD Attack Path Analysis
- [ADRecon](https://github.com/adrecon/ADRecon) - AD Reconnaissance
- [Purple Knight](https://www.purple-knight.com/) - AD Security Assessment

### Microsoft Tools
- [Microsoft Security Assessment Tool](https://learn.microsoft.com/en-us/assessments/)
- [Attack Surface Analyzer](https://github.com/microsoft/AttackSurfaceAnalyzer)
- [Azure AD Security Assessment](https://learn.microsoft.com/en-us/azure/active-directory/fundamentals/security-operations-introduction)

