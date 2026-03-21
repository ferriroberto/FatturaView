# 🔒 Risoluzione Errore Execution Policy PowerShell

## ❌ Errore

```
.\SignExecutable.ps1 : L'esecuzione di script è disabilitata nel sistema in uso.
PSSecurityException: UnauthorizedAccess
```

## 📋 Causa

Windows blocca l'esecuzione di script PowerShell per motivi di sicurezza. È necessario autorizzare l'esecuzione.

---

## ✅ SOLUZIONI

### 1️⃣ Bypass Temporaneo (RACCOMANDATO - Più Sicuro)

Esegui lo script con bypass una tantum **senza modificare** la policy di sistema:

```powershell
# Metodo A - Bypass nella riga di comando
powershell.exe -ExecutionPolicy Bypass -File .\SignExecutable.ps1

# Metodo B - Con parametri
powershell.exe -ExecutionPolicy Bypass -File .\SignExecutable.ps1 -CertPath "C:\Cert\MyCert.pfx" -CertPassword "pwd"
```

**Vantaggi**:
- ✅ Non modifica le impostazioni di sistema
- ✅ Sicuro - solo questo script viene eseguito
- ✅ Non serve essere amministratore

---

### 2️⃣ Sbloccare il File Scaricato

Se lo script è stato scaricato da Internet, Windows lo marca come "bloccato":

```powershell
# Sblocca il file
Unblock-File -Path .\SignExecutable.ps1

# Poi eseguilo normalmente
.\SignExecutable.ps1
```

**Verifica se il file è bloccato**:
```powershell
Get-Item .\SignExecutable.ps1 | Select-Object Name, @{Name="Blocked";Expression={$_.Attributes -band [System.IO.FileAttributes]::ReadOnly}}
```

---

### 3️⃣ Cambiare Policy per Utente Corrente (Permanente)

Cambia la policy solo per il tuo utente (non serve amministratore):

```powershell
# Abilita per l'utente corrente
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Conferma con 'Y' (Yes)
```

**Policy Impostate**:
- `RemoteSigned` = Gli script locali possono essere eseguiti, quelli scaricati devono essere firmati

**Verifica**:
```powershell
Get-ExecutionPolicy -List
```

Output atteso:
```
Scope          ExecutionPolicy
-----          ---------------
MachinePolicy  Undefined
UserPolicy     Undefined
Process        Undefined
CurrentUser    RemoteSigned
LocalMachine   Undefined
```

---

### 4️⃣ Cambiare Policy per Sistema (Serve Amministratore)

**⚠️ Sconsigliato** - Apre la sicurezza per tutti gli utenti:

```powershell
# Apri PowerShell come Amministratore
# Esegui:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```

---

## 🎯 RACCOMANDAZIONE

Per **FatturaView** usa il **Metodo 1** (Bypass):

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SignExecutable.ps1
```

È il più sicuro e non richiede modifiche permanenti al sistema.

---

## 📖 Execution Policy Disponibili

| Policy | Descrizione | Sicurezza |
|--------|-------------|-----------|
| **Restricted** | Nessuno script può essere eseguito (default) | ⭐⭐⭐⭐⭐ |
| **AllSigned** | Solo script firmati da publisher trusted | ⭐⭐⭐⭐ |
| **RemoteSigned** | Script locali OK, remoti devono essere firmati | ⭐⭐⭐ |
| **Unrestricted** | Tutti gli script OK, avviso per quelli remoti | ⭐⭐ |
| **Bypass** | Nessun blocco, nessun avviso | ⭐ |

---

## 🔍 Verificare Policy Corrente

```powershell
# Policy effettiva
Get-ExecutionPolicy

# Policy per tutti gli scope
Get-ExecutionPolicy -List
```

---

## 🐛 Troubleshooting

### Errore: "Set-ExecutionPolicy non riconosciuto"

**Soluzione**: Sei in PowerShell Core invece di Windows PowerShell. Usa:
```powershell
# Apri Windows PowerShell (non PowerShell Core)
# Cerca: "Windows PowerShell" nel menu Start
```

### Errore: "Accesso negato" quando cambio policy

**Soluzione**: Apri PowerShell come Amministratore:
1. Cerca "PowerShell" nel menu Start
2. Tasto destro → **"Esegui come amministratore"**
3. Riprova il comando

### Lo script si blocca su "Inserisci password"

**Soluzione**: Specifica la password come parametro:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SignExecutable.ps1 -CertPassword "tua_password"
```

---

## 🔐 Sicurezza

### Perché Esiste Questa Protezione?

PowerShell può eseguire codice potente sul sistema. La policy impedisce l'esecuzione accidentale di script malevoli.

### È Sicuro Usare Bypass?

✅ **SÌ**, se:
- Conosci la fonte dello script
- Hai verificato il contenuto
- Lo usi solo temporaneamente

❌ **NO**, se:
- Script da fonte sconosciuta
- Non hai verificato il codice
- Lo usi come policy permanente

### Verifica Integrità Script

Prima di usare Bypass, verifica che lo script non sia stato modificato:

```powershell
# Calcola hash SHA256
Get-FileHash .\SignExecutable.ps1 -Algorithm SHA256
```

**Hash originale**:
```
Algorithm: SHA256
Hash: [verrà calcolato dopo la creazione finale]
```

---

## 📝 Automatizzare Firma in Visual Studio

Per evitare di eseguire manualmente lo script, aggiungilo come **Post-Build Event**:

1. **Visual Studio** → Proprietà Progetto → **Build Events**
2. **Post-Build Event** → **Command Line**:
   ```
   powershell.exe -ExecutionPolicy Bypass -File "$(ProjectDir)SignExecutable.ps1" -ExePath "$(TargetPath)"
   ```

Ora l'eseguibile viene firmato automaticamente ad ogni build!

---

## 🎯 Quick Reference

```powershell
# ✅ MODO RACCOMANDATO (una tantum, sicuro)
powershell.exe -ExecutionPolicy Bypass -File .\SignExecutable.ps1

# ✅ MODO ALTERNATIVO (sblocca file)
Unblock-File -Path .\SignExecutable.ps1
.\SignExecutable.ps1

# ⚠️ MODO PERMANENTE (solo se necessario)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 🔍 VERIFICA POLICY CORRENTE
Get-ExecutionPolicy
```

---

## 📚 Risorse Aggiuntive

- **Microsoft Docs**: https://docs.microsoft.com/powershell/module/microsoft.powershell.core/about/about_execution_policies
- **Sicurezza PowerShell**: https://docs.microsoft.com/powershell/scripting/security
- **Set-ExecutionPolicy**: https://docs.microsoft.com/powershell/module/microsoft.powershell.security/set-executionpolicy

---

## ✅ Dopo la Risoluzione

Una volta risolto, procedi con la firma dell'eseguibile:

1. ✅ Execution policy risolta
2. ✅ Esegui: `powershell.exe -ExecutionPolicy Bypass -File .\SignExecutable.ps1`
3. ✅ Lo script firmerà FatturaView.exe
4. ✅ Avviso Foxit risolto!

---

**FatturaView** v1.0.0  
© 2026 Roberto Ferri
