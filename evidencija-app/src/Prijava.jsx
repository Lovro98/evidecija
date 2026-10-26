import React, { useEffect, useRef, useState } from "react";
import { supabase } from "./supabase.js";
import { POSITIONS } from "./candidates.js";

// Javna prijavnica: /prijava (bez prijave u aplikaciju).
// Samo ime, telefon i pozicija; regruter se kandidatu javlja telefonom, Viberom ili WhatsAppom.
// Oznaka izvora iz linka (/prijava?z=fb-praha) sprema se uz prijavu da se vidi iz koje grupe je došao.

const C = {
  bg: "#F4F5F1", card: "#FFFFFF", ink: "#1C2521", sub: "#5C6862", line: "#E1E5DF",
  green: "#0E6B4F", greenSoft: "#E3F0EA", red: "#B3402F", redSoft: "#F9E8E4",
};

const T = {
  UK: {
    title: "Робота в готелях Праги",
    lead: "Залиште контакт — ми вам зателефонуємо або напишемо у Viber / WhatsApp.",
    name: "Ім'я та прізвище", phone: "Номер телефону", positions: "Яка робота вас цікавить?", positionsHint: "Можна вибрати кілька",
    consent: "Погоджуюся, щоб мені зателефонували або написали щодо роботи.",
    send: "Надіслати", sending: "Надсилаю…",
    errName: "Вкажіть ім'я та прізвище.", errPhone: "Вкажіть правильний номер телефону.", errPos: "Виберіть хоча б одну роботу.",
    errConsent: "Поставте галочку, що погоджуєтеся на дзвінок.", errSend: "Не вдалося надіслати. Перевірте інтернет і спробуйте ще раз.",
    doneTitle: "Дякуємо!", doneText: "Ми зв'яжемося з вами найближчим часом.",
  },
  CS: {
    title: "Práce v hotelech v Praze",
    lead: "Nechte nám kontakt, zavoláme vám nebo napíšeme přes Viber / WhatsApp.",
    name: "Jméno a příjmení", phone: "Telefonní číslo", positions: "O jakou práci máte zájem?", positionsHint: "Můžete vybrat více možností",
    consent: "Souhlasím, že mi zavoláte nebo napíšete ohledně práce.",
    send: "Odeslat", sending: "Odesílám…",
    errName: "Vyplňte jméno a příjmení.", errPhone: "Vyplňte platné telefonní číslo.", errPos: "Vyberte alespoň jednu práci.",
    errConsent: "Zaškrtněte souhlas s kontaktováním.", errSend: "Odeslání se nepovedlo. Zkontrolujte internet a zkuste to znovu.",
    doneTitle: "Děkujeme!", doneText: "Brzy se vám ozveme.",
  },
};

const inputStyle = { width: "100%", boxSizing: "border-box", border: `1.5px solid ${C.line}`, borderRadius: 12, padding: "13px 14px",
  fontSize: 16, fontFamily: "inherit", color: C.ink, background: "#fff", outline: "none" };
const label = { display: "block", fontWeight: 700, fontSize: 15.5, marginBottom: 8 };

export default function Prijava() {
  const params = new URLSearchParams(window.location.search);
  const [lang, setLang] = useState((params.get("l") || "").toUpperCase() === "CS" ? "CS" : "UK");
  const [name, setName] = useState("");
  const [phone, setPhone] = useState("");
  const [positions, setPositions] = useState([]);
  const [consent, setConsent] = useState(false);
  const [website, setWebsite] = useState(""); // skriveno polje za botove
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);
  const openedAt = useRef(Date.now());
  const t = T[lang];
  const sourceTag = (params.get("z") || "").slice(0, 60);

  useEffect(() => { document.title = t.title; document.documentElement.lang = lang === "UK" ? "uk" : "cs"; }, [lang, t.title]);

  const toggle = (id) => setPositions((p) => (p.includes(id) ? p.filter((x) => x !== id) : [...p, id]));

  const submit = async (e) => {
    e.preventDefault();
    setErr("");
    if (name.trim().length < 2) return setErr(t.errName);
    if (phone.replace(/[^0-9]/g, "").length < 9) return setErr(t.errPhone);
    if (positions.length === 0) return setErr(t.errPos);
    if (!consent) return setErr(t.errConsent);
    if (website || Date.now() - openedAt.current < 3000) { setDone(true); return; }
    setBusy(true);
    const { error } = await supabase.rpc("submit_application", {
      p_full_name: name, p_phone: phone, p_messenger: "", p_positions: positions,
      p_experience: "", p_residence: "", p_city: "", p_start_from: "",
      p_heard_from: "", p_referrer: "", p_source_tag: sourceTag, p_lang: lang,
    });
    setBusy(false);
    if (error) return setErr(t.errSend);
    setDone(true);
  };

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.ink, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif" }}>
      <div style={{ maxWidth: 520, margin: "0 auto", padding: "16px 16px 48px" }}>
        <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 14 }}>
          <div style={{ display: "flex", gap: 4, background: "#fff", border: `1px solid ${C.line}`, borderRadius: 999, padding: 3 }}>
            {[["UK", "🇺🇦 UA"], ["CS", "🇨🇿 CZ"]].map(([id, l]) => (
              <button key={id} type="button" onClick={() => setLang(id)}
                style={{ border: "none", borderRadius: 999, padding: "6px 11px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit",
                  background: lang === id ? C.green : "transparent", color: lang === id ? "#fff" : C.sub }}>{l}</button>
            ))}
          </div>
        </div>

        {done ? (
          <div style={{ background: C.card, border: `1px solid ${C.line}`, borderRadius: 18, padding: "36px 22px", textAlign: "center" }}>
            <div style={{ width: 56, height: 56, borderRadius: 999, background: C.greenSoft, color: C.green, fontSize: 30, fontWeight: 800,
              display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>✓</div>
            <h1 style={{ fontSize: 24, margin: "0 0 8px" }}>{t.doneTitle}</h1>
            <p style={{ color: C.sub, fontSize: 16, lineHeight: 1.5, margin: 0 }}>{t.doneText}</p>
          </div>
        ) : (
          <form onSubmit={submit} noValidate>
            <h1 style={{ fontSize: 27, lineHeight: 1.2, margin: "0 0 8px" }}>{t.title}</h1>
            <p style={{ color: C.sub, fontSize: 15.5, lineHeight: 1.5, margin: "0 0 18px" }}>{t.lead}</p>

            <div style={{ background: C.card, border: `1px solid ${C.line}`, borderRadius: 18, padding: "20px 16px" }}>
              <div style={{ marginBottom: 18 }}>
                <label style={label} htmlFor="p-name">{t.name}</label>
                <input id="p-name" style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} autoComplete="name" maxLength={120} />
              </div>
              <div style={{ marginBottom: 18 }}>
                <label style={label} htmlFor="p-phone">{t.phone}</label>
                <input id="p-phone" style={inputStyle} value={phone} onChange={(e) => setPhone(e.target.value)} type="tel" inputMode="tel"
                  autoComplete="tel" placeholder="+420 …" maxLength={30} />
              </div>
              <div style={{ marginBottom: 18 }}>
                <div style={{ ...label, marginBottom: 2 }}>{t.positions}</div>
                <div style={{ color: C.sub, fontSize: 13, marginBottom: 10 }}>{t.positionsHint}</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  {POSITIONS.map((o) => {
                    const on = positions.includes(o.id);
                    return (
                      <button key={o.id} type="button" onClick={() => toggle(o.id)} aria-pressed={on}
                        style={{ border: `1.5px solid ${on ? C.green : C.line}`, background: on ? C.greenSoft : "#fff", color: on ? C.green : C.ink,
                          borderRadius: 999, padding: "10px 14px", fontSize: 15, fontWeight: on ? 700 : 500, cursor: "pointer", fontFamily: "inherit" }}>
                        {on ? "✓ " : ""}{o[lang]}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div aria-hidden style={{ position: "absolute", left: -9999, width: 1, height: 1, overflow: "hidden" }}>
                <input tabIndex={-1} autoComplete="off" name="website" value={website} onChange={(e) => setWebsite(e.target.value)} />
              </div>

              <label style={{ display: "flex", gap: 10, alignItems: "flex-start", fontSize: 14, color: C.sub, lineHeight: 1.45, cursor: "pointer" }}>
                <input type="checkbox" checked={consent} onChange={(e) => setConsent(e.target.checked)} style={{ width: 20, height: 20, marginTop: 0, flexShrink: 0 }} />
                <span>{t.consent}</span>
              </label>
            </div>

            {err && <div role="alert" style={{ background: C.redSoft, color: C.red, borderRadius: 12, padding: "11px 13px", fontSize: 14.5, fontWeight: 600, marginTop: 14 }}>{err}</div>}

            <button type="submit" disabled={busy}
              style={{ width: "100%", marginTop: 16, background: C.green, color: "#fff", border: "none", borderRadius: 14, padding: "15px 16px",
                fontSize: 17, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.7 : 1, fontFamily: "inherit" }}>
              {busy ? t.sending : t.send}
            </button>
          </form>
        )}
      </div>
    </div>
  );
}
