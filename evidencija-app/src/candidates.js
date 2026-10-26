// Zajedničke oznake za prijavnicu (/prijava) i admin panel Kandidati.
// U bazi se spremaju samo ključevi (npr. "hsk"), a natpisi se prevode ovdje.

export const POSITIONS = [
  { id: "hsk", UK: "Покоївка (housekeeping)", CS: "Pokojská (HSK)", HR: "Sobarica (HSK)", focus: true },
  { id: "cook", UK: "Кухар / кухарка", CS: "Kuchař / kuchařka", HR: "Kuhar / kuharica", focus: true },
  { id: "kitchen_help", UK: "Помічник кухаря", CS: "Pomocná síla v kuchyni", HR: "Pomoćni radnik u kuhinji", focus: true },
  { id: "dishwasher", UK: "Мийник посуду", CS: "Myčka nádobí", HR: "Perač suđa", focus: true },
  { id: "public_cleaning", UK: "Прибирання (хол, коридори)", CS: "Úklid veřejných prostor", HR: "Čišćenje javnih prostora" },
  { id: "waiter", UK: "Офіціант / офіціантка", CS: "Číšník / servírka", HR: "Konobar / konobarica" },
  { id: "bartender", UK: "Бармен", CS: "Barman", HR: "Barmen" },
  { id: "reception", UK: "Рецепція", CS: "Recepční", HR: "Recepcija" },
  { id: "maintenance", UK: "Технічне обслуговування", CS: "Údržba", HR: "Održavanje" },
  { id: "other", UK: "Інше", CS: "Jiné", HR: "Ostalo" },
];

export const RESIDENCE = [
  { id: "temp_protection", UK: "Тимчасовий захист", CS: "Dočasná ochrana", HR: "Privremena zaštita" },
  { id: "visa", UK: "Віза / дозвіл на проживання", CS: "Vízum / povolení k pobytu", HR: "Viza / boravišna dozvola" },
  { id: "eu", UK: "Громадянство ЄС", CS: "Občan EU", HR: "Državljanin EU" },
  { id: "other", UK: "Інше / не знаю", CS: "Jiné / nevím", HR: "Ostalo / ne zna" },
];

export const MESSENGERS = [
  { id: "viber", UK: "Viber", CS: "Viber", HR: "Viber" },
  { id: "telegram", UK: "Telegram", CS: "Telegram", HR: "Telegram" },
  { id: "whatsapp", UK: "WhatsApp", CS: "WhatsApp", HR: "WhatsApp" },
  { id: "phone", UK: "Тільки дзвінок", CS: "Jen telefon", HR: "Samo poziv" },
];

export const START = [
  { id: "now", UK: "Одразу", CS: "Ihned", HR: "Odmah" },
  { id: "week", UK: "Протягом тижня", CS: "Do týdne", HR: "Unutar tjedan dana" },
  { id: "month", UK: "Протягом місяця", CS: "Do měsíce", HR: "Unutar mjesec dana" },
  { id: "later", UK: "Пізніше", CS: "Později", HR: "Kasnije" },
];

export const HEARD = [
  { id: "fb", UK: "Facebook-група", CS: "Skupina na Facebooku", HR: "FB grupa" },
  { id: "viber", UK: "Viber-група", CS: "Skupina na Viberu", HR: "Viber grupa" },
  { id: "telegram", UK: "Telegram", CS: "Telegram", HR: "Telegram" },
  { id: "friend", UK: "Від знайомих / колег", CS: "Od známých / kolegů", HR: "Preporuka (poznanik / kolega)" },
  { id: "portal", UK: "Сайт з вакансіями", CS: "Pracovní portál", HR: "Portal za posao" },
  { id: "other", UK: "Інше", CS: "Jiné", HR: "Ostalo" },
];

export const CAND_STATUS = [
  { id: "novi", label: "Novi", color: "#B97F1E", bg: "#FBF1DC" },
  { id: "kontaktiran", label: "Kontaktiran", color: "#2B5D8A", bg: "#E6EEF5" },
  { id: "razgovor", label: "Razgovor", color: "#6B4BA8", bg: "#EFEAF8" },
  { id: "poceo", label: "Počeo raditi", color: "#0E6B4F", bg: "#E3F0EA" },
  { id: "nije_prikladan", label: "Nije prikladan", color: "#5C6862", bg: "#EEF0EC" },
  { id: "nedostupan", label: "Nedostupan", color: "#B3402F", bg: "#F9E8E4" },
];

export const labelOf = (list, id, lang = "HR") => list.find((x) => x.id === id)?.[lang] || id || "";
