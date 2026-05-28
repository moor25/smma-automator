import { useState } from "react";

const APPS_SCRIPT_CODE = `// =============================================
//  SMMA Meeting Automator – Google Apps Script
//  script.google.com → Új projekt → Bemásolni
// =============================================

// ⚙️ ITT ÁLLÍTSD BE A SAJÁT ADATAIDAT:
const YOUR_NAME = "Kovács Anna";          // A te neved
const YOUR_COMPANY = "Ügynökség neve";    // Cégnév
const TIMEZONE = "Europe/Budapest";

// Az e-mail leírás (az időpont automatikusan változik)
function getDescription(name, dateStr, meetLink) {
  return \`Szia \${name}!

Köszönöm az idődet a telefonhíváson. Megerősítem az egyeztetett konzultációnkat.

📅 Időpont: \${dateStr}
🎥 Google Meet link: \${meetLink}

Az agendánk:
• Jelenlegi marketing helyzeted rövid áttekintése
• Konkrét növekedési lehetőségek azonosítása
• Következő lépések megbeszélése

Ha bármilyen kérdésed van addig, csak válaszolj erre az e-mailre.

Hamarosan találkozunk!
\${YOUR_NAME}
\${YOUR_COMPANY}\`;
}

// -----------------------------------------------

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { email, name, datetime } = data;
    const meetingDate = new Date(datetime);

    // 1. Google Calendar esemény + Meet link
    const calEvent = createEventWithMeet(email, name, meetingDate);
    const meetLink = getMeetLink(calEvent);

    // 2. Visszaigazoló e-mail az ügyfélnek
    const dateStr = Utilities.formatDate(meetingDate, TIMEZONE, "yyyy. MMMM dd. (EEEE) HH:mm");
    GmailApp.sendEmail(
      email,
      "Visszaigazolás – Konzultáció | " + YOUR_COMPANY,
      getDescription(name, dateStr, meetLink)
    );

    // 3. Emlékeztetők ütemezése
    scheduleReminders(email, name, meetingDate, meetLink);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, meetLink }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error(err);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createEventWithMeet(clientEmail, clientName, startTime) {
  const endTime = new Date(startTime.getTime() + 3600000);
  return Calendar.Events.insert({
    summary: "SMMA Konzultáció – " + clientName,
    start: { dateTime: startTime.toISOString(), timeZone: TIMEZONE },
    end:   { dateTime: endTime.toISOString(),   timeZone: TIMEZONE },
    attendees: [{ email: clientEmail }],
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: "hangoutsMeet" }
      }
    }
  }, "primary", { conferenceDataVersion: 1 });
}

function getMeetLink(calEvent) {
  try { return calEvent.conferenceData.entryPoints[0].uri; }
  catch (e) { return "(meet link nem elérhető)"; }
}

function scheduleReminders(email, name, meetingDate, meetLink) {
  const now = new Date();
  const diffHours = (meetingDate - now) / 3600000;
  const reminders = [];

  if (diffHours > 48)  reminders.push({ offset: 48 * 60, type: "2 nap" });
  if (diffHours > 24)  reminders.push({ offset: 24 * 60, type: "1 nap" });
  if (diffHours > 1)   reminders.push({ offset: 60,      type: "1 óra" });
  if (diffHours > 0.5) reminders.push({ offset: 30,      type: "30 perc" });

  const props = PropertiesService.getScriptProperties();
  const existing = JSON.parse(props.getProperty("smma_reminders") || "[]");

  for (const r of reminders) {
    existing.push({
      email, name, meetLink,
      meetingDate: meetingDate.toISOString(),
      sendAt: new Date(meetingDate.getTime() - r.offset * 60000).toISOString(),
      type: r.type,
      sent: false
    });
  }
  props.setProperty("smma_reminders", JSON.stringify(existing));
}

// ▶ Ezt futtatsd el EGYSZER manuálisan a beállításhoz!
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("checkReminders")
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log("✅ Trigger beállítva: checkReminders minden 10 percben");
}

// Automatikusan fut minden 10 percben
function checkReminders() {
  const props = PropertiesService.getScriptProperties();
  const reminders = JSON.parse(props.getProperty("smma_reminders") || "[]");
  const now = new Date();

  const updated = reminders.map(r => {
    if (!r.sent && new Date(r.sendAt) <= now) {
      sendReminderEmail(r);
      return { ...r, sent: true };
    }
    return r;
  });

  const cutoff = new Date(Date.now() - 2 * 24 * 3600000);
  const cleaned = updated.filter(r => new Date(r.meetingDate) > cutoff);
  props.setProperty("smma_reminders", JSON.stringify(cleaned));
}

function sendReminderEmail(r) {
  const dateStr = Utilities.formatDate(new Date(r.meetingDate), TIMEZONE, "HH:mm");
  const subjects = {
    "2 nap":   "📅 Holnapután konzultáció – emlékeztető",
    "1 nap":   "📅 Holnap konzultáció – emlékeztető",
    "1 óra":   "⏰ 1 óra múlva kezdünk!",
    "30 perc": "🔔 30 perc múlva kezdünk!"
  };
  const bodies = {
    "2 nap":   \`Szia \${r.name}!\\n\\nEmlékeztetlek, hogy holnapután konzultációnk van \${dateStr}-kor.\\n\\n🎥 Meet link: \${r.meetLink}\\n\\nVárunk!\`,
    "1 nap":   \`Szia \${r.name}!\\n\\nEmlékeztetlek, hogy holnap konzultációnk van \${dateStr}-kor.\\n\\n🎥 Meet link: \${r.meetLink}\\n\\nVárunk!\`,
    "1 óra":   \`Szia \${r.name}!\\n\\n1 óra múlva kezdünk! \${dateStr}-kor várunk a Meet-en.\\n\\n🎥 \${r.meetLink}\\n\\nHamarosan találkozunk!\`,
    "30 perc": \`Szia \${r.name}!\\n\\n30 perc múlva kezdünk! Kattints a linkre és csatlakozz:\\n\\n🎥 \${r.meetLink}\\n\\nVárunk!\`
  };
  GmailApp.sendEmail(r.email, subjects[r.type] || "Emlékeztető", bodies[r.type] || "");
  Logger.log("📧 Emlékeztető elküldve: " + r.email + " (" + r.type + ")");
}`;

const STEPS = [
  {
    num: "1",
    title: "Nyisd meg a Google Apps Script oldalt",
    desc: "Menj a script.google.com oldalra, és kattints az \"Új projekt\" gombra.",
    link: { label: "script.google.com →", href: "https://script.google.com" }
  },
  {
    num: "2",
    title: "Illeszd be a kódot",
    desc: "Töröld az alapértelmezett tartalmat, majd másold be az alábbi kódot. A tetején állítsd be a neved és a cégedet.",
    code: true
  },
  {
    num: "3",
    title: "Mentés és engedélyezés",
    desc: "Nyomd meg a Mentés gombot (Ctrl+S), majd kattints a ▶ Futtatás gombra a setupTrigger függvényen. Google engedélyt kér – fogadd el mindent."
  },
  {
    num: "4",
    title: "Deploy (közzététel)",
    desc: "Deploy → New deployment → Web app. \"Execute as: Me\" | \"Who has access: Anyone\". Kattints a Deploy gombra, majd másold az URL-t."
  },
  {
    num: "5",
    title: "Illeszd be az URL-t lent",
    desc: "Másold be a kapott URL-t az alábbi mezőbe. Megmarad oldalfrissítés után is. Kész, minden automatikus!",
    urlInput: true
  }
];

const USERS = {
  coreadmin: { password: "kukac2005", role: "admin" },
  coreuser:  { password: "coreuser2026", role: "user" },
};

function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [pw, setPw] = useState("");
  const [error, setError] = useState(false);

  const handleLogin = () => {
    const user = USERS[username.trim().toLowerCase()];
    if (user && pw === user.password) {
      sessionStorage.setItem("smma_auth", user.role);
      onLogin(user.role);
    } else {
      setError(true);
      setTimeout(() => setError(false), 1500);
    }
  };

  const fieldStyle = (hasError) => ({
    width: "100%", background: "#0c0c14",
    border: `1px solid ${hasError ? "#ef4444" : "#1e1e2e"}`,
    borderRadius: 10, padding: "12px 14px", color: "#e2e8f0",
    fontSize: 14, outline: "none", marginBottom: 12,
    transition: "border-color 0.2s", boxSizing: "border-box",
  });

  return (
    <div style={{
      minHeight: "100vh", background: "#0c0c14", display: "flex",
      alignItems: "center", justifyContent: "center", padding: 16,
      fontFamily: "'Inter', 'Segoe UI', sans-serif",
    }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&display=swap');`}</style>
      <div style={{
        background: "#13131a", border: "1px solid #1e1e2e", borderRadius: 20,
        padding: "40px 32px", width: "100%", maxWidth: 360, textAlign: "center",
      }}>
        <div style={{
          width: 48, height: 48, borderRadius: "50%", background: "#6366f120",
          border: "1px solid #6366f140", display: "flex", alignItems: "center",
          justifyContent: "center", margin: "0 auto 20px", fontSize: 22,
        }}>🔒</div>
        <h2 style={{
          fontFamily: "'Syne', sans-serif", fontSize: 22, fontWeight: 800,
          color: "#e2e8f0", margin: "0 0 6px",
        }}>CORE Marketing</h2>
        <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 28px" }}>
          Belépés a meeting rendszerbe
        </p>
        <input
          type="text"
          placeholder="Felhasználónév"
          value={username}
          onChange={e => setUsername(e.target.value)}
          autoFocus
          style={fieldStyle(error)}
        />
        <input
          type="password"
          placeholder="Jelszó"
          value={pw}
          onChange={e => setPw(e.target.value)}
          onKeyDown={e => e.key === "Enter" && handleLogin()}
          style={fieldStyle(error)}
        />
        {error && (
          <p style={{ color: "#ef4444", fontSize: 12, margin: "0 0 12px" }}>
            Helytelen felhasználónév vagy jelszó
          </p>
        )}
        <button
          onClick={handleLogin}
          style={{
            width: "100%", background: "#6366f1", color: "#fff", border: "none",
            borderRadius: 10, padding: "12px", fontSize: 14, fontWeight: 600,
            cursor: "pointer",
          }}
        >
          Belépés →
        </button>
      </div>
    </div>
  );
}

export default function App() {
  const [role, setRole] = useState(() => sessionStorage.getItem("smma_auth") || "");
  const [tab, setTab] = useState("schedule");
  const [scriptUrl, setScriptUrl] = useState(() => localStorage.getItem("smma_script_url") || "");
  const [form, setForm] = useState({ email: "", name: "", datetime: "" });
  const [status, setStatus] = useState("idle");
  const [copied, setCopied] = useState(false);

  const meetingDate = form.datetime ? new Date(form.datetime) : null;
  const diffHours = meetingDate ? (meetingDate - new Date()) / 3600000 : 0;

  const reminders = (() => {
    if (!meetingDate || diffHours <= 0) return [];
    const r = [];
    if (diffHours > 48) r.push("2 nappal előtte");
    if (diffHours > 24) r.push("1 nappal előtte");
    if (diffHours > 1)  r.push("1 órával előtte");
    if (diffHours > 0.5) r.push("30 perccel előtte");
    return r;
  })();

  const handleSubmit = () => {
    if (!scriptUrl || !form.email || !form.datetime) return;
    setStatus("loading");

    const params = new URLSearchParams({
      email: form.email,
      name: form.name,
      datetime: form.datetime,
    });
    const popup = window.open(
      `${scriptUrl}?${params}`,
      "_blank",
      "width=1,height=1,left=0,top=0"
    );
    setTimeout(() => popup?.close(), 5000);

    setTimeout(() => {
      setStatus("success");
      setForm({ email: "", name: "", datetime: "" });
    }, 2000);
  };

  const copyCode = () => {
    navigator.clipboard.writeText(APPS_SCRIPT_CODE);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  if (!role) return <LoginScreen onLogin={(r) => setRole(r)} />;
  const isAdmin = role === "admin";

  return (
    <div style={{
      minHeight: "100vh",
      background: "#0c0c14",
      color: "#e2e8f0",
      fontFamily: "'Inter', 'Segoe UI', sans-serif",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      padding: "36px 16px 60px",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Syne:wght@700;800&display=swap');
        * { box-sizing: border-box; }
        input, textarea { outline: none; }
        input:focus { border-color: #818cf8 !important; box-shadow: 0 0 0 3px #818cf822 !important; }
        .tab:hover { color: #e2e8f0 !important; }
        .submit-btn:hover:not(:disabled) { background: #6366f1 !important; transform: translateY(-1px); }
        .copy-btn:hover { background: #1e1e2e !important; color: #818cf8 !important; }
        .step-card { transition: border-color 0.2s; }
        .step-card:hover { border-color: #6366f144 !important; }
        pre { overflow-x: auto; }
        pre::-webkit-scrollbar { height: 6px; }
        pre::-webkit-scrollbar-track { background: #0c0c14; }
        pre::-webkit-scrollbar-thumb { background: #2d2d44; border-radius: 3px; }
        @keyframes fadeUp { from { opacity:0; transform: translateY(12px); } to { opacity:1; transform:translateY(0); } }
        .fu { animation: fadeUp 0.35s ease forwards; }
        .fu1 { animation-delay: 0s; opacity:0; }
        .fu2 { animation-delay: .07s; opacity:0; }
        .fu3 { animation-delay: .14s; opacity:0; }
        .fu4 { animation-delay: .21s; opacity:0; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .spinner { animation: spin 0.8s linear infinite; }
        @keyframes pop { 0% { transform:scale(0.8); opacity:0; } 60% { transform:scale(1.08); } 100% { transform:scale(1); opacity:1; } }
        .pop { animation: pop 0.4s ease forwards; }
        @keyframes blink { 0%,100% { opacity:1; } 50% { opacity:0.3; } }
        .blink { animation: blink 2s ease infinite; }
      `}</style>

      {/* Header */}
      <div style={{ textAlign: "center", marginBottom: 32 }} className="fu fu1">
        <div style={{
          display: "inline-flex", alignItems: "center", gap: 8,
          background: "#6366f115", border: "1px solid #6366f133",
          borderRadius: 100, padding: "5px 14px", marginBottom: 14,
          fontSize: 12, color: "#a5b4fc", fontWeight: 500, letterSpacing: "0.06em",
          textTransform: "uppercase",
        }}>
          <span style={{ width: 6, height: 6, borderRadius: "50%", background: "#22d3a5", display: "inline-block" }} className="blink" />
          SMMA Automátor
        </div>
        <h1 style={{
          fontFamily: "'Syne', sans-serif",
          fontSize: "clamp(28px, 6vw, 46px)",
          fontWeight: 800, margin: "0 0 8px", lineHeight: 1.1,
          background: "linear-gradient(135deg, #fff 30%, #818cf8)",
          WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent",
        }}>
          Cold Call → Meeting
        </h1>
        <p style={{ color: "#64748b", fontSize: 14, margin: 0 }}>
          Megadod az adatokat — minden más automatikus
        </p>
      </div>

      {/* Tabs — csak adminnak */}
      {isAdmin && (
        <div style={{
          display: "flex", background: "#13131a", border: "1px solid #1e1e2e",
          borderRadius: 12, padding: 4, marginBottom: 28, gap: 2,
        }} className="fu fu2">
          {[["schedule", "📅 Ütemezés"], ["setup", "⚙️ Beállítás"]].map(([key, label]) => (
            <button key={key} className="tab" onClick={() => setTab(key)} style={{
              background: tab === key ? "#6366f1" : "transparent",
              color: tab === key ? "#fff" : "#64748b",
              border: "none", borderRadius: 9, padding: "9px 22px",
              fontSize: 14, fontWeight: 500, cursor: "pointer", transition: "all 0.2s",
            }}>{label}</button>
          ))}
        </div>
      )}

      <div style={{ width: "100%", maxWidth: 540 }}>

        {/* SCHEDULE TAB */}
        {tab === "schedule" && (
          <div>
            {status === "success" ? (
              <div style={{
                background: "#13131a", border: "1px solid #22d3a533",
                borderRadius: 20, padding: "40px 28px", textAlign: "center",
              }} className="pop">
                <div style={{ fontSize: 52, marginBottom: 16 }}>✅</div>
                <h2 style={{ fontFamily: "'Syne',sans-serif", fontSize: 24, fontWeight: 700, margin: "0 0 8px", color: "#22d3a5" }}>
                  Kész!
                </h2>
                <p style={{ color: "#94a3b8", fontSize: 14, margin: "0 0 8px" }}>
                  Az ügyfél megkapta a visszaigazolást a Google Meet linkkel együtt.
                </p>
                <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 28px" }}>
                  Az emlékeztetők automatikusan ki fognak menni.
                </p>
                <button onClick={() => setStatus("idle")} style={{
                  background: "#6366f1", color: "#fff", border: "none",
                  borderRadius: 10, padding: "12px 28px", fontSize: 14,
                  fontWeight: 600, cursor: "pointer",
                }}>
                  ← Új meeting
                </button>
              </div>
            ) : (
              <div style={{ background: "#13131a", border: "1px solid #1e1e2e", borderRadius: 20, padding: "28px 24px" }} className="fu fu1">

                {!scriptUrl && isAdmin && (
                  <div style={{ background: "#f59e0b11", border: "1px solid #f59e0b33", borderRadius: 10, padding: "10px 14px", marginBottom: 20, fontSize: 13, color: "#fcd34d", cursor: "pointer" }}
                    onClick={() => setTab("setup")}>
                    ⚠️ Még nincs URL beállítva — kattints ide a Beállítás fülhöz
                  </div>
                )}

                <div style={{ height: 1, background: "#1e1e2e", margin: "0 0 20px" }} />

                <div style={{ marginBottom: 16 }} className="fu fu2">
                  <label style={labelStyle}>Ügyfél e-mail *</label>
                  <input type="email" placeholder="pelda@ceg.hu" value={form.email}
                    onChange={e => setForm(f => ({ ...f, email: e.target.value }))}
                    style={inputStyle} />
                </div>

                <div style={{ marginBottom: 16 }} className="fu fu2">
                  <label style={labelStyle}>Ügyfél neve *</label>
                  <input type="text" placeholder="Kovács Péter" value={form.name}
                    onChange={e => setForm(f => ({ ...f, name: e.target.value }))}
                    style={inputStyle} />
                </div>

                <div style={{ marginBottom: 24 }} className="fu fu3">
                  <label style={labelStyle}>Egyeztetett időpont *</label>
                  <input type="datetime-local" value={form.datetime}
                    onChange={e => setForm(f => ({ ...f, datetime: e.target.value }))}
                    style={{ ...inputStyle, colorScheme: "dark" }} />
                </div>

                {reminders.length > 0 && (
                  <div style={{
                    background: "#6366f10d", border: "1px solid #6366f133",
                    borderRadius: 12, padding: "14px 16px", marginBottom: 20,
                  }} className="fu fu3">
                    <p style={{ fontSize: 11, color: "#a5b4fc", fontWeight: 600, margin: "0 0 10px", textTransform: "uppercase", letterSpacing: "0.07em" }}>
                      🔔 Automatikus emlékeztetők
                    </p>
                    {reminders.map((r, i) => (
                      <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "#e2e8f0", marginBottom: i < reminders.length - 1 ? 6 : 0 }}>
                        <span style={{ color: "#818cf8" }}>→</span> {r}
                      </div>
                    ))}
                  </div>
                )}

                {status === "error" && (
                  <div style={{ background: "#ef444415", border: "1px solid #ef444433", borderRadius: 10, padding: "10px 14px", marginBottom: 16, fontSize: 13, color: "#fca5a5" }}>
                    ❌ Hiba történt. Ellenőrizd az Apps Script URL-t és a beállításokat.
                  </div>
                )}

                <button
                  className="submit-btn"
                  onClick={handleSubmit}
                  disabled={status === "loading" || !form.email || !form.name || !form.datetime || !scriptUrl}
                  style={{
                    width: "100%", background: "#4f46e5",
                    color: "#fff", border: "none", borderRadius: 12,
                    padding: "14px", fontSize: 15, fontWeight: 600,
                    cursor: "pointer", transition: "all 0.2s",
                    opacity: (!form.email || !form.name || !form.datetime || !scriptUrl) ? 0.4 : 1,
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                  }}
                >
                  {status === "loading" ? (
                    <>
                      <span style={{ width: 16, height: 16, border: "2px solid #ffffff44", borderTopColor: "#fff", borderRadius: "50%" }} className="spinner" />
                      Küldés...
                    </>
                  ) : "✦ Meeting létrehozása"}
                </button>

                <p style={{ textAlign: "center", color: "#334155", fontSize: 11, margin: "14px 0 0" }}>
                  → Naptárbejegyzés + visszaigazoló e-mail + összes emlékeztető
                </p>
              </div>
            )}
          </div>
        )}

        {/* SETUP TAB — csak admin */}
        {tab === "setup" && isAdmin && (
          <div>
            <div style={{ background: "#13131a", border: "1px solid #1e1e2e", borderRadius: 20, padding: "24px" }} className="fu fu1">
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontSize: 18, fontWeight: 700, margin: "0 0 6px" }}>
                Egyszeri beállítás (~10 perc)
              </h2>
              <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 24px" }}>
                Csak egyszer kell elvégezni, utána minden automatikus.
              </p>

              {STEPS.map((s, i) => (
                <div key={i} className="step-card" style={{
                  display: "flex", gap: 16, marginBottom: 20,
                  paddingBottom: i < STEPS.length - 1 ? 20 : 0,
                  borderBottom: i < STEPS.length - 1 ? "1px solid #1e1e2e" : "none",
                }}>
                  <div style={{
                    width: 32, height: 32, borderRadius: "50%", flexShrink: 0,
                    background: "#6366f1", display: "flex", alignItems: "center",
                    justifyContent: "center", fontWeight: 700, fontSize: 14,
                    fontFamily: "'Syne',sans-serif",
                  }}>{s.num}</div>
                  <div style={{ flex: 1 }}>
                    <p style={{ fontWeight: 600, fontSize: 14, margin: "0 0 4px", paddingTop: 6 }}>{s.title}</p>
                    <p style={{ color: "#64748b", fontSize: 13, margin: 0, lineHeight: 1.6 }}>{s.desc}</p>
                    {s.link && (
                      <a href={s.link.href} target="_blank" rel="noreferrer" style={{
                        color: "#818cf8", fontSize: 13, fontWeight: 500,
                        display: "inline-block", marginTop: 6,
                      }}>{s.link.label}</a>
                    )}
                    {s.urlInput && (
                      <div style={{ marginTop: 10 }}>
                        <input
                          type="url"
                          placeholder="https://script.google.com/macros/s/.../exec"
                          value={scriptUrl}
                          onChange={e => {
                            setScriptUrl(e.target.value);
                            localStorage.setItem("smma_script_url", e.target.value);
                          }}
                          style={{ ...inputStyle, fontSize: 12, borderColor: scriptUrl ? "#22d3a533" : "#f59e0b55" }}
                        />
                        {scriptUrl && (
                          <p style={{ color: "#22d3a5", fontSize: 11, margin: "6px 0 0" }}>✓ URL mentve</p>
                        )}
                      </div>
                    )}
                    {s.code && (
                      <div style={{ marginTop: 12, position: "relative" }}>
                        <div style={{
                          background: "#0c0c14", border: "1px solid #1e1e2e",
                          borderRadius: 10, padding: "12px 14px", maxHeight: 160,
                          overflow: "hidden", position: "relative",
                        }}>
                          <pre style={{ margin: 0, fontSize: 11, color: "#64748b", lineHeight: 1.5, fontFamily: "monospace" }}>
                            {APPS_SCRIPT_CODE.slice(0, 300)}...
                          </pre>
                          <div style={{
                            position: "absolute", bottom: 0, left: 0, right: 0, height: 60,
                            background: "linear-gradient(transparent, #0c0c14)",
                          }} />
                        </div>
                        <button className="copy-btn" onClick={copyCode} style={{
                          marginTop: 8, width: "100%",
                          background: "#1e1e2e", border: "1px solid #2d2d44",
                          borderRadius: 8, padding: "10px", fontSize: 13,
                          color: copied ? "#22d3a5" : "#94a3b8",
                          fontWeight: 500, cursor: "pointer", transition: "all 0.2s",
                        }}>
                          {copied ? "✓ Kód másolva!" : "📋 Teljes kód másolása"}
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>

            <div style={{
              background: "#13131a", border: "1px solid #1e1e2e",
              borderRadius: 16, padding: "20px 24px", marginTop: 16,
            }} className="fu fu2">
              <p style={{ fontSize: 12, color: "#6366f1", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em", margin: "0 0 14px" }}>
                ✦ Mit csinál automatikusan?
              </p>
              {[
                ["📅", "Google Calendar esemény", "Google Meet linkkel, az ügyfél meghívót kap"],
                ["✉️", "Visszaigazoló e-mail", "Azonnal kimegy a saját leírásod szövegével"],
                ["🔔", "2 nappal előtte", "Ha több mint 2 nap van hátra"],
                ["🔔", "1 nappal előtte", "Ha több mint 24 óra van hátra"],
                ["⏰", "1 órával előtte", "Ha aznap van"],
                ["🔔", "30 perccel előtte", "Ha aznap van"],
              ].map(([icon, title, desc], i) => (
                <div key={i} style={{ display: "flex", gap: 12, marginBottom: 10 }}>
                  <span style={{ fontSize: 16, flexShrink: 0 }}>{icon}</span>
                  <div>
                    <span style={{ fontSize: 13, fontWeight: 500 }}>{title}</span>
                    <span style={{ color: "#64748b", fontSize: 12 }}> — {desc}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const labelStyle = {
  fontSize: 11, fontWeight: 600, color: "#64748b",
  letterSpacing: "0.08em", textTransform: "uppercase",
  display: "block", marginBottom: 7,
};

const inputStyle = {
  width: "100%", background: "#0c0c14",
  border: "1px solid #1e1e2e", borderRadius: 10,
  padding: "11px 14px", color: "#e2e8f0", fontSize: 14,
  transition: "all 0.2s",
};
