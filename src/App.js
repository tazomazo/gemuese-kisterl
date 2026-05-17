import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

const ADMIN_PASSWORD = "admin123";
const WEEK_LABEL = "KW 16 · 2026";

const INITIAL_PRODUCTS = [
  { name: "Österr. Bio-Rhabarber", price: 5.0, unit: "500g", category: "Frühlingsneuheiten", special: true },
  { name: "Österreichischer Spargel, weiß", price: 7.0, unit: "500g", category: "Frühlingsneuheiten", special: true },
  { name: "Österreichischer Spargel, grün", price: 7.0, unit: "500g", category: "Frühlingsneuheiten", special: true },
  { name: "Waldviertler Bärlauch", price: 2.4, unit: "100g", category: "Frühlingsneuheiten", special: true },
  { name: "Italienische Heurige", price: 3.6, unit: "Kilo", category: "Kartoffeln & Wurzeln", special: false },
  { name: "Waldviertler Süßkartoffel", price: 4.8, unit: "Kilo", category: "Kartoffeln & Wurzeln", special: false },
  { name: "Marchfelder Sigma, Erdäpfel", price: 3.6, unit: "Kilo", category: "Kartoffeln & Wurzeln", special: false },
  { name: "Waldviertler 'Linzer Rose'", price: 2.8, unit: "Kilo", category: "Kartoffeln & Wurzeln", special: false },
  { name: "Ital. Bio-Avocado", price: 1.0, unit: "Stück", category: "Obst & Früchte", special: false },
  { name: "Ital. Bio-Grapefruit Rose", price: 1.3, unit: "Stück", category: "Obst & Früchte", special: false },
  { name: "Ital. Bio-Blut-Orangen", price: 3.9, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Steirische Bio-Erdbeeren", price: 6.0, unit: "500g", category: "Obst & Früchte", special: true },
  { name: "Steirische Bio-Himbeeren", price: 4.5, unit: "250g", category: "Obst & Früchte", special: false },
  { name: "Steirische Bio-Heidelbeeren", price: 4.5, unit: "250g", category: "Obst & Früchte", special: false },
  { name: "Steirische Ribisel", price: 6.0, unit: "500g", category: "Obst & Früchte", special: false },
  { name: "Ital. Weintrauben, kernlos", price: 2.8, unit: "500g", category: "Obst & Früchte", special: false },
  { name: "Ital. Bio Kiwi", price: 1.0, unit: "Stück", category: "Obst & Früchte", special: false },
  { name: "Ital. Bio-Orangen", price: 3.8, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Ital. Bio-Mandarinen", price: 3.9, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Ital. Bio-Zitronen", price: 0.6, unit: "Stück", category: "Obst & Früchte", special: false },
  { name: "Steir. Birnen Williams", price: 3.8, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Steir. Bio-Rubinetten Äpfel", price: 3.6, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Steir. Bio-Elstar Äpfel", price: 3.6, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Steir. Bio-Gala Äpfel", price: 3.6, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Fairtrade Bananen aus Kolumbien", price: 2.6, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Bio Bananen aus Peru", price: 3.2, unit: "Kilo", category: "Obst & Früchte", special: false },
  { name: "Wr. Gärtner Pak Choi", price: 2.4, unit: "500g", category: "Blatt & Salat", special: false },
  { name: "Wr. Gärtner Spinat", price: 3.0, unit: "500g", category: "Blatt & Salat", special: false },
  { name: "Wr.Gärtner Mangold", price: 4.8, unit: "Kilo", category: "Blatt & Salat", special: false },
  { name: "Wr. Gärtner Fisolen", price: 4.0, unit: "500g", category: "Blatt & Salat", special: false },
  { name: "Wr. Gärtner Kohlsprossen", price: 5.0, unit: "500g", category: "Blatt & Salat", special: false },
  { name: "Wr.Gärtner Salat", price: 1.3, unit: "Stück", category: "Blatt & Salat", special: false },
  { name: "Wr.Gärtner Bummerl (Eisbergsalat)", price: 1.5, unit: "Stück", category: "Blatt & Salat", special: false },
  { name: "Wr.Gärtner Vogerlsalat", price: 1.8, unit: "100g", category: "Blatt & Salat", special: false },
  { name: "Wr.Gärtner Rucola", price: 1.8, unit: "100g", category: "Blatt & Salat", special: false },
  { name: "Österr. Jägersalat (Chinakohl)", price: 2.8, unit: "Kilo", category: "Blatt & Salat", special: false },
  { name: "Wr.Gärtner Chicorée", price: 3.0, unit: "500g", category: "Blatt & Salat", special: false },
  { name: "Wr.Gärtner Gurken", price: 1.3, unit: "Stück", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Radieschen", price: 1.3, unit: "Bund", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Jungzwiebel", price: 1.4, unit: "Bund", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Cocktailparadeiser", price: 3.4, unit: "500g", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Aromaparadeiser", price: 2.8, unit: "500g", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Fleischparadeiser", price: 2.4, unit: "500g", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Paprika rot/gelb/grün", price: 1.3, unit: "Stück", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Fenchel", price: 2.4, unit: "500g", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Melanzani", price: 4.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Porree", price: 3.9, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Kohlrabi", price: 1.3, unit: "Stück", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Kohl", price: 2.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Marchfelder Kraut, weiß/rot", price: 2.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Zucchini", price: 4.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Brokkoli", price: 4.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Wr.Gärtner Karfiol", price: 3.6, unit: "Stück", category: "Gemüse", special: false },
  { name: "Marchfelder Rote Rüben", price: 3.0, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Suppengemüse, klein", price: 1.4, unit: "ca. 500g", category: "Gemüse", special: false },
  { name: "Suppengemüse, groß", price: 2.8, unit: "ca. 1kg", category: "Gemüse", special: false },
  { name: "Marchfelder Zeller", price: 2.4, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Marchfelder Gelbe Rüben", price: 2.4, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Marchfelder Petersilwurzeln", price: 2.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Marchfelder Karotten", price: 2.0, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Marchfelder Bundkarotten", price: 2.8, unit: "Bund", category: "Gemüse", special: false },
  { name: "Steir. Bio-Kürbis Butternuß", price: 3.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Steir. Bio-Kürbis Hokkaido", price: 3.8, unit: "Kilo", category: "Gemüse", special: false },
  { name: "Marchfelder Mai-/Herbst-Rüben", price: 2.0, unit: "Bund (3 Stück)", category: "Gemüse", special: false },
  { name: "Kärntner Champignons", price: 2.9, unit: "500g", category: "Pilze", special: false },
  { name: "Raasdorfer Kräuterseitlinge", price: 12.0, unit: "500g", category: "Pilze", special: false },
  { name: "NÖ Austernpilze", price: 3.5, unit: "250g", category: "Pilze", special: false },
  { name: "Laaer Bio-Zwiebel (heurige Ernte)", price: 2.4, unit: "Kilo", category: "Zwiebeln & Knoblauch", special: false },
  { name: "Laaer Bio-Zwiebel rot", price: 2.4, unit: "Kilo", category: "Zwiebeln & Knoblauch", special: false },
  { name: "NÖ Schalottenzwiebel", price: 1.2, unit: "250g", category: "Zwiebeln & Knoblauch", special: false },
  { name: "Bgdl. Knoblauch", price: 1.5, unit: "Stück (Happerl)", category: "Zwiebeln & Knoblauch", special: false },
  { name: "Wr.Gärtner Bio-Gartenkresse", price: 1.1, unit: "Schachterl", category: "Kräuter", special: false },
  { name: "Wr.Gärtner Basilikum", price: 2.5, unit: "50g", category: "Kräuter", special: false },
  { name: "Wr.Gärtner Koriander (grün)", price: 2.5, unit: "50g", category: "Kräuter", special: false },
  { name: "Wr.Gärtner Rosmarin", price: 2.5, unit: "50g", category: "Kräuter", special: false },
  { name: "Wr.Gärtner Thymian", price: 2.5, unit: "50g", category: "Kräuter", special: false },
  { name: "Petersilie", price: 1.2, unit: "Bund", category: "Kräuter", special: false },
  { name: "Schnittlauch", price: 1.2, unit: "Bund", category: "Kräuter", special: false },
  { name: "Dille", price: 1.5, unit: "Bund", category: "Kräuter", special: false },
  { name: "Eferdinger Ingwer", price: 1.0, unit: "100g", category: "Kräuter", special: false },
  { name: "Steir. Nüsse, ausgelöst", price: 4.5, unit: "250g", category: "Nüsse & Extras", special: false },
  { name: "Ital. Bio-Erdnüsse", price: 2.6, unit: "250g", category: "Nüsse & Extras", special: false },
  { name: "GRANOLA (200g Glas + €1 Pfand)", price: 5.9, unit: "Glas", category: "Nüsse & Extras", special: true },
  { name: "Steir. Kren, gerieben", price: 3.6, unit: "Glas (100g)", category: "Nüsse & Extras", special: false },
  { name: "Steir. Kren", price: 1.0, unit: "100g", category: "Nüsse & Extras", special: false },
  { name: "Mühlv. Bio-Apfel-Birnen Essig", price: 4.9, unit: "500ml", category: "Öle & Essige", special: false },
  { name: "Mühlv. Bio-Apfelessig (Eichenfass)", price: 4.9, unit: "500ml", category: "Öle & Essige", special: false },
  { name: "Mühlv. Bio-Birnenbalsam Essig", price: 5.3, unit: "250ml", category: "Öle & Essige", special: false },
  { name: "Steir. Kernöl", price: 13.8, unit: "500ml", category: "Öle & Essige", special: false },
  { name: "Steir. Himbeernektar (40%)", price: 4.8, unit: "Liter", category: "Säfte & Honig", special: false },
  { name: "Steir. Heidelbeernektar (40%)", price: 4.8, unit: "Liter", category: "Säfte & Honig", special: false },
  { name: "Steir. Marillennektar (40%)", price: 4.8, unit: "Liter", category: "Säfte & Honig", special: false },
  { name: "Mühlv. Bio-Apfel-Johannisbeere", price: 3.9, unit: "Liter", category: "Säfte & Honig", special: false },
  { name: "Mühlv. Bio-Birnensaft", price: 3.5, unit: "Liter", category: "Säfte & Honig", special: false },
  { name: "Mühlv. Bio-Apfelsaft, naturtrüb", price: 2.9, unit: "Liter", category: "Säfte & Honig", special: false },
  { name: "Blütenhonig aus Wien Mauer", price: 10.0, unit: "500g", category: "Säfte & Honig", special: false },
  { name: "OÖ Bio-Freilandeier, Größe XL", price: 0.5, unit: "Stück", category: "Eier", special: false },
];

const categoryIcons = {
  "Frühlingsneuheiten": "🌱", "Kartoffeln & Wurzeln": "🥔",
  "Obst & Früchte": "🍎", "Blatt & Salat": "🥬", Gemüse: "🥦",
  Pilze: "🍄", "Zwiebeln & Knoblauch": "🧅", Kräuter: "🌿",
  "Nüsse & Extras": "🌰", "Öle & Essige": "🫙", "Säfte & Honig": "🍯",
  Eier: "🥚", "Alle Produkte": "📦",
};

// ── Supabase helpers ──────────────────────────────────────────────────────────

async function dbLoadProducts() {
  const { data, error } = await supabase.from("products").select("*, categories(name)").order("id");
  if (error) throw error;
  return data.map((p) => ({ ...p, category: p.categories?.name ?? "" }));
}
async function dbSeedProducts(products) {
  const { error } = await supabase.from("products").insert(products);
  if (error) throw error;
}
async function dbReplaceProducts(products) {
  await supabase.from("products").delete().neq("id", 0);
  const { error } = await supabase.from("products").insert(products);
  if (error) throw error;
}
async function dbUpdateProduct(id, fields) {
  const { error } = await supabase.from("products").update(fields).eq("id", id);
  if (error) throw error;
}
async function dbDeleteProduct(id) {
  const { error } = await supabase.from("products").delete().eq("id", id);
  if (error) throw error;
}
async function dbAddProduct(product) {
  const { data, error } = await supabase.from("products").insert(product).select().single();
  if (error) throw error;
  return data;
}

async function dbLoadUsers() {
  const { data, error } = await supabase.from("users").select("id, name, password").order("name");
  if (error) throw error;
  return data;
}
async function dbCreateUser(name, password) {
  const { data, error } = await supabase
    .from("users").insert({ name, password }).select().single();
  if (error) throw error;
  return data;
}

async function dbLoadOrder(userId) {
  const { data } = await supabase.from("orders").select("cart").eq("user_id", userId).maybeSingle();
  return data ? data.cart : {};
}
async function dbSaveOrder(userId, cart) {
  const { error } = await supabase.from("orders")
    .upsert({ user_id: userId, cart, updated_at: new Date().toISOString() }, { onConflict: "user_id" });
  if (error) throw error;
}
async function dbLoadCategories() {
  const { data, error } = await supabase.from("categories").select("*").order("sort_order").order("name");
  if (error) throw error;
  return data;
}
async function dbSeedCategories(cats) {
  const { error } = await supabase.from("categories").insert(cats);
  if (error) throw error;
}
async function dbAddCategory(name) {
  const { data, error } = await supabase.from("categories").insert({ name, sort_order: 999 }).select().single();
  if (error) throw error;
  return data;
}
async function dbRenameCategory(id, newName) {
  const { error } = await supabase.from("categories").update({ name: newName }).eq("id", id);
  if (error) throw error;
}
async function dbDeleteCategory(id) {
  const { error } = await supabase.from("categories").delete().eq("id", id);
  if (error) throw error;
}

async function dbLoadAllOrders() {
  const { data, error } = await supabase
    .from("orders")
    .select("cart, updated_at, users(id, name)")
    .order("updated_at", { ascending: false });
  if (error) throw error;
  return data || [];
}

function guessProductCategory(name) {
  const n = name.toLowerCase();
  if (n.includes("spargel") || n.includes("rhabarber") || n.includes("bärlauch")) return "Frühlingsneuheiten";
  if (n.includes("kartoffel") || n.includes("erdäpfel") || n.includes("süßkartoffel") || n.includes("heurige") || n.includes("linzer rose")) return "Kartoffeln & Wurzeln";
  if (n.includes("apfel") || n.includes("birne") || n.includes("beere") || n.includes("traube") || n.includes("kiwi") ||
      n.includes("avocado") || n.includes("banane") || n.includes("orange") || n.includes("mandarine") ||
      n.includes("zitrone") || n.includes("grapefruit") || n.includes("marill") || n.includes("ribisel") ||
      n.includes("maroni")) return "Obst & Früchte";
  if (n.includes("champignon") || n.includes("seitling") || n.includes("austernpilz") || n.includes("pilz")) return "Pilze";
  if (n.includes("zwiebel") || n.includes("knoblauch") || n.includes("schalott")) return "Zwiebeln & Knoblauch";
  if (n.includes("basilikum") || n.includes("koriander") || n.includes("rosmarin") || n.includes("thymian") ||
      n.includes("petersilie") || n.includes("schnittlauch") || n.includes("dille") || n.includes("ingwer") ||
      n.includes("kresse")) return "Kräuter";
  if (n.includes("nuss") || n.includes("nüsse") || n.includes("erdnuss") || n.includes("granola") || n.includes("kren")) return "Nüsse & Extras";
  if (n.includes("essig") || n.includes("kernöl") || (n.includes("öl") && !n.includes("kohlrabi"))) return "Öle & Essige";
  if (n.includes("nektar") || n.includes("saft") || n.includes("honig")) return "Säfte & Honig";
  if (n.includes("eier") || n.includes("freilandei") || (n.includes(" ei") && n.includes("bio"))) return "Eier";
  if (n.includes("salat") || n.includes("spinat") || n.includes("mangold") || n.includes("fisolen") ||
      n.includes("kohlsprossen") || n.includes("pak choi") || n.includes("rucola") || n.includes("vogerlsalat") ||
      n.includes("chicorée") || n.includes("jägersalat") || n.includes("chinakohl")) return "Blatt & Salat";
  return "Gemüse";
}

function parseXlsx(file, existingProducts = [], categories = []) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });

        // Prefer sheet named "Gemüse", otherwise first sheet
        const sheetName =
          wb.SheetNames.find((n) =>
            n.toLowerCase().includes("gem") || n.toLowerCase().includes("produkt")
          ) || wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
        const products = [];
        const seenNames = new Set();

        for (const row of rows) {
          // Dynamically find the first column with a proper product name string
          let nameIdx = -1;
          let name = null;
          for (let i = 0; i < Math.min(row.length, 4); i++) {
            const cell = row[i];
            if (cell && typeof cell === "string" && cell.trim().length > 2) {
              nameIdx = i;
              name = cell.trim();
              break;
            }
          }
          if (!name || nameIdx < 0) continue;

          // Price is the next column, unit the one after
          const rawPrice = row[nameIdx + 1];
          const rawUnit  = row[nameIdx + 2];

          // Parse price robustly (handle "3,60" or plain numbers)
          let price = null;
          if (typeof rawPrice === "number" && rawPrice > 0) {
            price = rawPrice;
          } else if (typeof rawPrice === "string") {
            const cleaned = rawPrice.replace(",", ".").replace(/[^0-9.]/g, "");
            const p = parseFloat(cleaned);
            if (!isNaN(p) && p > 0) price = p;
          }
          if (!price) continue;

          // Skip duplicate names (continuation rows from merged cells)
          const key = name.toLowerCase();
          if (seenNames.has(key)) continue;
          seenNames.add(key);

          products.push({
            name,
            price,
            unit: rawUnit ? String(rawUnit).trim() : "",
            category_id: (() => {
              const catName = existingProducts.find((ep) => ep.name.toLowerCase() === name.toLowerCase())?.category ?? guessProductCategory(name);
              return categories.find((c) => c.name === catName)?.id ?? categories[0]?.id ?? null;
            })(),
            special: false,
          });
        }

        resolve(products);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

const NEWSLETTER_TEMPLATE_URL = "https://mailchi.mp/515a9f1d7d9f/andreas-gemsekisterl-woche-18344583?e=869a666e7b";
const NEWSLETTER_FONT = "'Open Sans', 'Helvetica Neue', Helvetica, Arial, sans-serif";

function escHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

async function generateNewsletterHtml(products) {
  const apiRes = await fetch(`/api/fetch-newsletter?url=${encodeURIComponent(NEWSLETTER_TEMPLATE_URL)}`);
  if (!apiRes.ok) throw new Error("Vorlage nicht ladbar");
  const html = await apiRes.text();

  const doc = new DOMParser().parseFromString(html, "text/html");

  // Find product ULs: contain <strong> and a € price
  const productUls = Array.from(doc.querySelectorAll("ul")).filter(
    (ul) => ul.textContent.includes("€") && ul.querySelector("strong")
  );
  if (productUls.length === 0) throw new Error("Produktbereich in Vorlage nicht gefunden");

  // Build one <li> per product
  const makeItem = (p) => {
    const price = p.price.toFixed(2).replace(".", ",");
    return (
      `<li style="color: #ffffff;"><p style="line-height: 1.5; mso-line-height-alt: 150%;">` +
      `<strong><span style="color:#ffffff;"><span style="font-family: ${NEWSLETTER_FONT}">${escHtml(p.name)}</span></span></strong>` +
      `<span style="color:#ffffff;"><span style="font-size: 14px"><span style="font-family: ${NEWSLETTER_FONT}"> € ${price}/${escHtml(p.unit)}</span></span></span>` +
      `</p></li>`
    );
  };

  // Replace first product UL with all products; remove the rest
  productUls[0].innerHTML = products.map(makeItem).join("\n");
  for (let i = 1; i < productUls.length; i++) productUls[i].remove();

  // Replace week label (KW XX) with current label
  const body = doc.body;
  body.innerHTML = body.innerHTML.replace(/KW\s*\d+\s*[·•-]?\s*\d{4}/g, WEEK_LABEL);

  return "<!DOCTYPE html>\n" + doc.documentElement.outerHTML;
}

// ── Root ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [screen, setScreen] = useState("landing");
  const [currentUser, setCurrentUser] = useState(null); // { id, name }
  const [isAdmin, setIsAdmin] = useState(false); // eslint-disable-line no-unused-vars
  const [products, setProducts] = useState([]);
  const [dbCategories, setDbCategories] = useState([]);
  const [dbReady, setDbReady] = useState(false);
  const [orders, setOrders] = useState([]);
  const [cart, setCart] = useState({});
  const [activeCategory, setActiveCategory] = useState("Alle");
  const [search, setSearch] = useState("");
  const [toast, setToast] = useState(null);
  const [uploadMsg, setUploadMsg] = useState("");
  const [loading, setLoading] = useState(false);
  const fileRef = useRef();

  useEffect(() => {
    (async () => {
      try {
        // 1. Load/seed categories first (products need category_id)
        let cats = await dbLoadCategories();
        if (cats.length === 0) {
          await dbSeedCategories(
            Object.keys(categoryIcons).filter((c) => c !== "Alle Produkte").map((name, i) => ({ name, sort_order: i }))
          );
          cats = await dbLoadCategories();
        }
        setDbCategories(cats);
        const catMap = Object.fromEntries(cats.map((c) => [c.name, c.id]));

        // 2. Load/seed products
        const data = await dbLoadProducts();
        if (data.length === 0) {
          await dbSeedProducts(
            INITIAL_PRODUCTS.map((p) => ({ name: p.name, price: p.price, unit: p.unit, special: p.special, category_id: catMap[p.category] ?? cats[0]?.id }))
          );
          setProducts(await dbLoadProducts());
        } else {
          setProducts(data);
        }
      } catch (e) {
        console.error(e);
        setProducts(INITIAL_PRODUCTS.map((p, i) => ({ ...p, id: i + 1 })));
      }
      setDbReady(true);
    })();
  }, []);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  };

  const handleUserLogin = async (user) => {
    setLoading(true);
    const savedCart = await dbLoadOrder(user.id);
    setCurrentUser(user);
    setCart(savedCart);
    setIsAdmin(false);
    setScreen("shop");
    setLoading(false);
  };

  const handleAdminLogin = async () => {
    setLoading(true);
    const allOrders = await dbLoadAllOrders();
    setOrders(allOrders);
    setIsAdmin(true);
    setScreen("admin");
    setLoading(false);
  };

  const handleLogout = () => {
    setCurrentUser(null); setIsAdmin(false); setCart({});
    setScreen("landing"); setSearch(""); setActiveCategory("Alle");
  };

  const saveCart = async () => {
    if (!Object.values(cart).some((q) => q > 0)) { showToast("Warenkorb ist leer.", "warn"); return; }
    setLoading(true);
    try {
      await dbSaveOrder(currentUser.id, cart);
      showToast("Bestellung gespeichert! ✓");
    } catch { showToast("Fehler beim Speichern.", "warn"); }
    setLoading(false);
  };

  const updateQty = (id, delta) => setCart((prev) => {
    const next = Math.max(0, parseFloat(((prev[id] || 0) + delta).toFixed(2)));
    if (next === 0) { const c = { ...prev }; delete c[id]; return c; }
    return { ...prev, [id]: next };
  });

  const setQty = (id, val) => {
    const n = parseFloat(val);
    if (isNaN(n) || n < 0) return;
    setCart((prev) => {
      if (n === 0) { const c = { ...prev }; delete c[id]; return c; }
      return { ...prev, [id]: n };
    });
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    setLoading(true);
    try {
      const parsed = await parseXlsx(file, products, dbCategories);
      if (parsed.length === 0) { setUploadMsg("⚠ Keine Produkte gefunden."); setLoading(false); return; }
      await dbReplaceProducts(parsed);
      setProducts(await dbLoadProducts());
      setUploadMsg(`✓ ${parsed.length} Produkte importiert aus "${file.name}"`);
    } catch { setUploadMsg("⚠ Fehler beim Importieren."); }
    setLoading(false); e.target.value = "";
  };

  const categories = [...new Set(products.map((p) => p.category))];
  const filteredProducts = products.filter((p) => {
    const matchCat = activeCategory === "Alle" || p.category === activeCategory;
    return matchCat && p.name.toLowerCase().includes(search.toLowerCase());
  });
  const cartTotal = Object.entries(cart).reduce((sum, [id, qty]) => {
    const p = products.find((x) => x.id === parseInt(id));
    return sum + (p ? p.price * qty : 0);
  }, 0);
  const cartCount = Object.values(cart).reduce((a, b) => a + b, 0);

  if (!dbReady) return <Splash />;

  if (screen === "landing") return (
    <Landing
      onUserLogin={() => setScreen("login-user")}
      onAdminLogin={() => setScreen("login-admin")}
    />
  );
  if (screen === "login-user") return (
    <UserLoginScreen
      onLogin={handleUserLogin}
      onBack={() => setScreen("landing")}
      loading={loading}
    />
  );
  if (screen === "login-admin") return (
    <AdminLoginScreen
      onLogin={handleAdminLogin}
      onBack={() => setScreen("landing")}
      loading={loading}
    />
  );
  if (screen === "admin") return (
    <AdminPanel
      orders={orders} products={products}
      categories={dbCategories}
      onCategoriesChange={async () => setDbCategories(await dbLoadCategories())}
      onUploadClick={() => fileRef.current.click()}
      uploadMsg={uploadMsg} onLogout={handleLogout}
      onRefresh={async () => setOrders(await dbLoadAllOrders())}
      onReloadProducts={async () => setProducts(await dbLoadProducts())}
      loading={loading}
    >
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleFileUpload} />
    </AdminPanel>
  );
  return (
    <ShopScreen
      user={currentUser} products={filteredProducts} allProducts={products}
      categories={categories} activeCategory={activeCategory} onCategory={setActiveCategory}
      search={search} onSearch={setSearch}
      cart={cart} cartTotal={cartTotal} cartCount={cartCount}
      onQty={updateQty} onSetQty={setQty}
      onSave={saveCart} onLogout={handleLogout}
      toast={toast} loading={loading} weekLabel={WEEK_LABEL}
    />
  );
}

// ── Splash ────────────────────────────────────────────────────────────────────

function Splash() {
  return (
    <div style={{ minHeight: "100vh", background: "#f0fdf4", display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 16 }}>
      <div style={{ fontSize: 48 }}>🌿</div>
      <div style={{ color: "#166534", fontWeight: 600, fontSize: 16 }}>Verbinde mit Datenbank…</div>
    </div>
  );
}

// ── Landing ───────────────────────────────────────────────────────────────────

function Landing({ onUserLogin, onAdminLogin }) {
  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #f0fdf4 0%, #fefce8 100%)", display: "flex", flexDirection: "column" }}>
      <header style={{ padding: "2rem 2.5rem 1rem" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 32 }}>🌿</span>
          <div>
            <div style={{ fontWeight: 700, fontSize: 20, color: "#166534" }}>Andreas' Gemüsekisterl</div>
            <div style={{ fontSize: 12, color: "#4ade80", fontWeight: 500, fontStyle: "italic" }}>à la carte · {WEEK_LABEL}</div>
          </div>
        </div>
      </header>
      <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "2rem" }}>
        <div style={{ textAlign: "center", maxWidth: 560 }}>
          <div style={{ fontSize: 72, marginBottom: "1rem" }}>🥦</div>
          <h1 style={{ fontSize: 42, fontWeight: 800, color: "#14532d", letterSpacing: "-1.5px", marginBottom: "0.75rem", lineHeight: 1.1 }}>
            Frisch. Regional.<br />Wöchentlich.
          </h1>
          <p style={{ color: "#4b7c59", fontSize: 17, marginBottom: "2.5rem", lineHeight: 1.6 }}>
            Bestelle dein Gemüse bequem online — direkt vom Wiener Gärtner & ausgewählten Produzenten.
          </p>
          <div style={{ display: "flex", gap: 16, justifyContent: "center", flexWrap: "wrap" }}>
            <button onClick={onUserLogin} style={btnPrimary}>Jetzt bestellen</button>
            <button onClick={onAdminLogin} style={btnSecondary}>Administrator</button>
          </div>
        </div>
        <div style={{ marginTop: "4rem", display: "flex", gap: 24, flexWrap: "wrap", justifyContent: "center" }}>
          {[
            { icon: "🌱", label: "Saisonale Neuheiten", sub: "Rhabarber, Spargel & mehr" },
            { icon: "📦", label: "Wöchentliche Bestellung", sub: "Jederzeit anpassbar" },
            { icon: "🏡", label: "Aus Wien & Österreich", sub: "Direkt vom Erzeuger" },
          ].map((f) => (
            <div key={f.label} style={{ background: "white", borderRadius: 16, padding: "20px 24px", minWidth: 180, boxShadow: "0 2px 12px rgba(0,0,0,0.06)", textAlign: "center" }}>
              <div style={{ fontSize: 28, marginBottom: 8 }}>{f.icon}</div>
              <div style={{ fontWeight: 600, fontSize: 14, color: "#14532d", marginBottom: 4 }}>{f.label}</div>
              <div style={{ fontSize: 12, color: "#4b7c59" }}>{f.sub}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── User Login ────────────────────────────────────────────────────────────────

function UserLoginScreen({ onLogin, onBack, loading }) {
  const [users, setUsers] = useState([]);
  const [mode, setMode] = useState("select"); // "select" | "create"
  const [selectedUser, setSelectedUser] = useState(null);
  const [password, setPassword] = useState("");
  const [newName, setNewName] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [newPassword2, setNewPassword2] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    dbLoadUsers().then(setUsers).catch(console.error);
  }, []);

  const handleSelect = async () => {
    if (!selectedUser) { setError("Bitte einen Namen auswählen."); return; }
    if (selectedUser.password && selectedUser.password !== password) {
      setError("Falsches Passwort."); return;
    }
    setBusy(true);
    await onLogin(selectedUser);
    setBusy(false);
  };

  const handleCreate = async () => {
    if (!newName.trim()) { setError("Bitte einen Namen eingeben."); return; }
    if (users.find((u) => u.name.toLowerCase() === newName.trim().toLowerCase())) {
      setError("Dieser Name existiert bereits."); return;
    }
    if (newPassword && newPassword !== newPassword2) { setError("Passwörter stimmen nicht überein."); return; }
    setBusy(true);
    try {
      const user = await dbCreateUser(newName.trim(), newPassword);
      await onLogin(user);
    } catch { setError("Fehler beim Anlegen des Benutzers."); }
    setBusy(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f0fdf4", display: "flex", alignItems: "center", justifyContent: "center", padding: "1rem" }}>
      <div style={{ background: "white", borderRadius: 20, padding: "2.5rem", width: "100%", maxWidth: 420, boxShadow: "0 4px 32px rgba(0,0,0,0.08)" }}>
        <button onClick={onBack} style={backBtn}>← Zurück</button>
        <div style={{ fontSize: 36, marginBottom: "0.75rem" }}>🌿</div>
        <h2 style={{ fontWeight: 700, fontSize: 24, color: "#14532d", marginBottom: 6 }}>
          {mode === "select" ? "Anmelden" : "Neuer Benutzer"}
        </h2>
        <p style={{ color: "#4b7c59", fontSize: 14, marginBottom: "1.5rem" }}>
          {mode === "select"
            ? "Wähle deinen Namen aus der Liste oder lege einen neuen Benutzer an."
            : "Gib deinen Namen ein und optional ein Passwort."}
        </p>

        {/* Tab switcher */}
        <div style={{ display: "flex", gap: 6, marginBottom: "1.25rem", background: "#f0fdf4", borderRadius: 10, padding: 4 }}>
          {["select", "create"].map((m) => (
            <button key={m} onClick={() => { setMode(m); setError(""); setPassword(""); }}
              style={{ flex: 1, background: mode === m ? "white" : "transparent", border: "none", borderRadius: 8, padding: "8px", fontSize: 13, fontWeight: mode === m ? 700 : 400, color: mode === m ? "#166534" : "#4b7c59", cursor: "pointer" }}>
              {m === "select" ? "Anmelden" : "Neu anlegen"}
            </button>
          ))}
        </div>

        {mode === "select" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {users.length === 0 ? (
              <div style={{ textAlign: "center", color: "#4b7c59", fontSize: 14, padding: "1rem 0" }}>
                Noch keine Benutzer vorhanden.<br />Bitte lege zuerst einen neuen Benutzer an.
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 220, overflowY: "auto" }}>
                {users.map((u) => (
                  <button key={u.id} onClick={() => { setSelectedUser(u); setError(""); setPassword(""); }}
                    style={{ background: selectedUser?.id === u.id ? "#f0fdf4" : "white", border: "1.5px solid " + (selectedUser?.id === u.id ? "#86efac" : "#e8f5e9"), borderRadius: 10, padding: "10px 14px", textAlign: "left", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <span style={{ fontWeight: 600, fontSize: 14, color: "#14532d" }}>👤 {u.name}</span>
                    {u.password && <span style={{ fontSize: 11, color: "#4b7c59" }}>🔒 Passwort</span>}
                  </button>
                ))}
              </div>
            )}
            {selectedUser?.password && (
              <input type="password" placeholder="Passwort eingeben" value={password}
                onChange={(e) => setPassword(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleSelect()}
                style={inputStyle} />
            )}
            {error && <ErrorBox msg={error} />}
            <button onClick={handleSelect} disabled={busy || loading || !selectedUser}
              style={{ ...btnPrimary, opacity: busy || !selectedUser ? 0.6 : 1, marginTop: 4 }}>
              {busy ? "Laden…" : "Weiter zum Shop →"}
            </button>
          </div>
        )}

        {mode === "create" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <input type="text" placeholder="Dein Name *" value={newName}
              onChange={(e) => setNewName(e.target.value)} style={inputStyle} />
            <input type="password" placeholder="Passwort (optional)" value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)} style={inputStyle} />
            {newPassword && (
              <input type="password" placeholder="Passwort wiederholen" value={newPassword2}
                onChange={(e) => setNewPassword2(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleCreate()}
                style={inputStyle} />
            )}
            {error && <ErrorBox msg={error} />}
            <button onClick={handleCreate} disabled={busy || loading}
              style={{ ...btnPrimary, opacity: busy ? 0.6 : 1, marginTop: 4 }}>
              {busy ? "Anlegen…" : "Benutzer anlegen & anmelden →"}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Admin Login ───────────────────────────────────────────────────────────────

function AdminLoginScreen({ onLogin, onBack, loading }) {
  const [name, setName] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const handle = async () => {
    if (!name.trim()) { setError("Bitte einen Namen eingeben."); return; }
    if (pass !== ADMIN_PASSWORD) { setError("Falsches Passwort."); return; }
    setBusy(true);
    await onLogin();
    setBusy(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f0fdf4", display: "flex", alignItems: "center", justifyContent: "center", padding: "1rem" }}>
      <div style={{ background: "white", borderRadius: 20, padding: "2.5rem", width: "100%", maxWidth: 400, boxShadow: "0 4px 32px rgba(0,0,0,0.08)" }}>
        <button onClick={onBack} style={backBtn}>← Zurück</button>
        <div style={{ fontSize: 36, marginBottom: "0.75rem" }}>🌿</div>
        <h2 style={{ fontWeight: 700, fontSize: 24, color: "#14532d", marginBottom: 6 }}>Administrator</h2>
        <p style={{ color: "#4b7c59", fontSize: 14, marginBottom: "1.5rem" }}>Name und Passwort eingeben.</p>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <input type="text" placeholder="Name" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle} />
          <input type="password" placeholder="Passwort" value={pass}
            onChange={(e) => setPass(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handle()}
            style={inputStyle} />
          {error && <ErrorBox msg={error} />}
          <button onClick={handle} disabled={busy || loading}
            style={{ ...btnPrimary, opacity: busy ? 0.6 : 1, marginTop: 4 }}>
            {busy ? "Laden…" : "Als Admin anmelden →"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Shop ──────────────────────────────────────────────────────────────────────

const qBtn = {
  background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 6,
  width: 28, height: 28, fontSize: 16, cursor: "pointer", color: "#166534",
  display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, lineHeight: 1,
};

function ShopScreen({ user, products, allProducts, categories, activeCategory, onCategory, search, onSearch, cart, cartTotal, cartCount, onQty, onSetQty, onSave, onLogout, toast, loading, weekLabel }) {
  const [view, setView] = useState("shop");
  const specials = allProducts.filter((p) => p.special);
  const cartItems = allProducts.filter((p) => cart[p.id]);

  return (
    <div style={{ minHeight: "100vh", background: "#f8fafb" }}>
      {toast && (
        <div style={{ position: "fixed", top: 20, right: 20, zIndex: 999, background: toast.type === "success" ? "#16a34a" : "#ca8a04", color: "white", borderRadius: 12, padding: "12px 20px", fontWeight: 600, fontSize: 14, boxShadow: "0 4px 20px rgba(0,0,0,0.15)" }}>
          {toast.msg}
        </div>
      )}
      <header style={{ background: "white", borderBottom: "1px solid #e8f5e9", padding: "0 1.5rem", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 22 }}>🌿</span>
            <div>
              <span style={{ fontWeight: 700, color: "#166534", fontSize: 16 }}>Andreas' Gemüsekisterl</span>
              <span style={{ fontSize: 11, color: "#4b7c59", marginLeft: 6, fontStyle: "italic" }}>à la carte</span>
            </div>
          </div>
          <nav style={{ display: "flex", gap: 4 }}>
            {["shop", "cart"].map((v) => (
              <button key={v} onClick={() => setView(v)} style={{ background: view === v ? "#f0fdf4" : "none", color: view === v ? "#16a34a" : "#4b7c59", border: "none", borderRadius: 8, padding: "6px 14px", fontWeight: view === v ? 700 : 500, cursor: "pointer", fontSize: 14 }}>
                {v === "shop" ? "Shop" : `Warenkorb${cartCount > 0 ? ` (${cartCount})` : ""}`}
              </button>
            ))}
          </nav>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 13, color: "#4b7c59" }}>👋 {user?.name}</span>
            <button onClick={onLogout} style={{ background: "none", border: "1px solid #bbf7d0", color: "#4b7c59", borderRadius: 8, padding: "5px 12px", fontSize: 13, cursor: "pointer" }}>Abmelden</button>
          </div>
        </div>
      </header>

      {view === "shop" && (
        <div style={{ maxWidth: 1100, margin: "0 auto", padding: "1.5rem" }}>
          {specials.length > 0 && (
            <div style={{ background: "linear-gradient(135deg, #f0fdf4, #fefce8)", borderRadius: 16, padding: "1.25rem 1.5rem", marginBottom: "1.5rem", border: "1px solid #bbf7d0" }}>
              <div style={{ fontWeight: 700, fontSize: 16, color: "#166534", marginBottom: "0.75rem" }}>🌱 Frühlingsneuheiten · {weekLabel}</div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                {specials.map((p) => (
                  <div key={p.id} style={{ background: "white", borderRadius: 10, padding: "10px 14px", border: "1px solid #dcfce7", display: "flex", alignItems: "center", gap: 10 }}>
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: "#166534" }}>{p.name}</div>
                      <div style={{ fontSize: 12, color: "#4b7c59" }}>€ {p.price.toFixed(2)} / {p.unit}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <button onClick={() => onQty(p.id, -0.5)} style={qBtn}>−</button>
                      <span style={{ fontSize: 13, minWidth: 22, textAlign: "center", color: "#166534", fontWeight: 600 }}>{cart[p.id] || 0}</span>
                      <button onClick={() => onQty(p.id, 0.5)} style={qBtn}>+</button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
          <div style={{ display: "flex", gap: 12, marginBottom: "1rem" }}>
            <input type="text" placeholder="🔍 Produkt suchen…" value={search} onChange={(e) => onSearch(e.target.value)}
              style={{ flex: 1, border: "1.5px solid #bbf7d0", borderRadius: 10, padding: "9px 14px", fontSize: 14, outline: "none", color: "#14532d" }} />
          </div>
          <div style={{ display: "flex", gap: 6, marginBottom: "1.25rem", overflowX: "auto", paddingBottom: 4 }}>
            {["Alle", ...categories].map((cat) => (
              <button key={cat} onClick={() => onCategory(cat)} style={{ background: activeCategory === cat ? "#16a34a" : "white", color: activeCategory === cat ? "white" : "#4b7c59", border: "1px solid " + (activeCategory === cat ? "#16a34a" : "#bbf7d0"), borderRadius: 20, padding: "5px 14px", fontSize: 13, cursor: "pointer", whiteSpace: "nowrap", fontWeight: activeCategory === cat ? 600 : 400 }}>
                {categoryIcons[cat] || ""} {cat}
              </button>
            ))}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 12 }}>
            {products.map((p) => <ProductCard key={p.id} product={p} qty={cart[p.id] || 0} onQty={onQty} onSetQty={onSetQty} />)}
          </div>
          {products.length === 0 && <div style={{ textAlign: "center", color: "#4b7c59", padding: "3rem" }}>Keine Produkte gefunden.</div>}
        </div>
      )}

      {view === "cart" && (
        <div style={{ maxWidth: 700, margin: "2rem auto", padding: "0 1.5rem" }}>
          <h2 style={{ fontWeight: 700, fontSize: 22, color: "#14532d", marginBottom: "1.25rem" }}>Deine Bestellung</h2>
          {cartItems.length === 0 ? (
            <div style={{ textAlign: "center", padding: "3rem", color: "#4b7c59" }}>
              <div style={{ fontSize: 48, marginBottom: 12 }}>🛒</div>
              <div>Dein Warenkorb ist noch leer.</div>
            </div>
          ) : (
            <>
              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: "1.5rem" }}>
                {cartItems.map((p) => (
                  <div key={p.id} style={{ background: "white", borderRadius: 12, padding: "14px 18px", border: "1px solid #e8f5e9", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ fontWeight: 600, fontSize: 14, color: "#14532d" }}>{p.name}</div>
                      <div style={{ fontSize: 12, color: "#4b7c59" }}>€ {p.price.toFixed(2)} / {p.unit}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <button onClick={() => onQty(p.id, -0.5)} style={qBtn}>−</button>
                      <input type="number" min="0" step="0.5" value={cart[p.id]}
                        onChange={(e) => onSetQty(p.id, e.target.value)}
                        style={{ width: 52, border: "1px solid #bbf7d0", borderRadius: 6, padding: "4px 6px", fontSize: 13, textAlign: "center", color: "#166534", fontWeight: 600 }} />
                      <button onClick={() => onQty(p.id, 0.5)} style={qBtn}>+</button>
                      <span style={{ minWidth: 64, textAlign: "right", fontSize: 13, fontWeight: 700, color: "#166534" }}>€ {(p.price * cart[p.id]).toFixed(2)}</span>
                    </div>
                  </div>
                ))}
              </div>
              <div style={{ background: "#f0fdf4", borderRadius: 12, padding: "1rem 1.25rem", display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem", border: "1px solid #bbf7d0" }}>
                <span style={{ fontWeight: 700, color: "#14532d", fontSize: 16 }}>Gesamtbetrag</span>
                <span style={{ fontWeight: 800, color: "#166534", fontSize: 20 }}>€ {cartTotal.toFixed(2)}</span>
              </div>
              <button onClick={onSave} disabled={loading}
                style={{ width: "100%", background: "#16a34a", color: "white", border: "none", borderRadius: 12, padding: "15px", fontSize: 16, fontWeight: 700, cursor: "pointer", opacity: loading ? 0.7 : 1 }}>
                {loading ? "Speichern…" : "Bestellung speichern ✓"}
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
}

function ProductCard({ product, qty, onQty, onSetQty }) {
  return (
    <div style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: qty > 0 ? "2px solid #86efac" : "1px solid #e8f5e9", transition: "border 0.15s" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 14, fontWeight: 600, color: "#14532d", lineHeight: 1.3, marginBottom: 3 }}>{product.name}</div>
          <div style={{ fontSize: 12, color: "#4b7c59" }}>{product.unit}</div>
        </div>
        {product.special && <span style={{ background: "#fef9c3", color: "#854d0e", fontSize: 10, fontWeight: 700, borderRadius: 6, padding: "2px 6px", marginLeft: 8 }}>NEU</span>}
      </div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontWeight: 800, fontSize: 16, color: "#166534" }}>€ {product.price.toFixed(2)}</span>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <button onClick={() => onQty(product.id, -0.5)} style={qBtn}>−</button>
          <input type="number" min="0" step="0.5" value={qty || ""} placeholder="0"
            onChange={(e) => onSetQty(product.id, e.target.value)}
            style={{ width: 44, border: "1px solid #bbf7d0", borderRadius: 6, padding: "4px", fontSize: 13, textAlign: "center", color: "#166534", fontWeight: 600 }} />
          <button onClick={() => onQty(product.id, 0.5)} style={qBtn}>+</button>
        </div>
      </div>
    </div>
  );
}

// ── Admin Panel ───────────────────────────────────────────────────────────────

function AdminPanel({ orders, products, categories, onCategoriesChange, onUploadClick, uploadMsg, onLogout, onRefresh, onReloadProducts, loading, children }) {
  const [tab, setTab] = useState("orders");
  const [editingId, setEditingId] = useState(null);
  const [editFields, setEditFields] = useState({ name: "", price: "", unit: "", category_id: null });
  const [showNewForm, setShowNewForm] = useState(false);
  const [newProduct, setNewProduct] = useState({ name: "", price: "", unit: "", category_id: null });
  const [productSearch, setProductSearch] = useState("");
  const [exportLoading, setExportLoading] = useState(false);
  const [exportMsg, setExportMsg] = useState("");
  const [editingCatId, setEditingCatId] = useState(null);
  const [editCatName, setEditCatName] = useState("");
  const [showNewCatForm, setShowNewCatForm] = useState(false);
  const [newCatName, setNewCatName] = useState("");

  const handleAddCategory = async () => {
    if (!newCatName.trim()) return;
    try {
      await dbAddCategory(newCatName.trim());
      await onCategoriesChange();
      setShowNewCatForm(false);
      setNewCatName("");
    } catch (e) { console.error(e); }
  };

  const handleRenameCategory = async (cat) => {
    if (!editCatName.trim() || editCatName.trim() === cat.name) { setEditingCatId(null); return; }
    try {
      await dbRenameCategory(cat.id, editCatName.trim());
      await onCategoriesChange();
      setEditingCatId(null);
    } catch (e) { console.error(e); }
  };

  const handleDeleteCategory = async (cat) => {
    const count = products.filter((p) => p.category_id === cat.id).length;
    if (count > 0) { alert(`Kategorie wird von ${count} Produkt(en) verwendet. Die Produkte werden keiner Kategorie zugeordnet.`); } // eslint-disable-line no-alert
    try {
      await dbDeleteCategory(cat.id);
      await onCategoriesChange();
      await onReloadProducts();
    } catch (e) { console.error(e); }
  };

  const handleSaveProduct = async (id) => {
    const price = parseFloat(editFields.price);
    if (!editFields.name.trim() || isNaN(price) || price <= 0) return;
    try {
      await dbUpdateProduct(id, { name: editFields.name.trim(), price, unit: editFields.unit.trim(), category_id: editFields.category_id });
      await onReloadProducts();
      setEditingId(null);
    } catch (e) { console.error(e); }
  };

  const handleDeleteProduct = async (id) => {
    if (!window.confirm("Produkt wirklich löschen?")) return; // eslint-disable-line no-restricted-globals
    try {
      await dbDeleteProduct(id);
      await onReloadProducts();
    } catch (e) { console.error(e); }
  };

  const handleNewsletterExport = async () => {
    setExportLoading(true);
    setExportMsg("Generiere Newsletter…");
    try {
      const html = await generateNewsletterHtml(products);
      const blob = new Blob([html], { type: "text/html" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `newsletter-${new Date().toISOString().slice(0, 10)}.html`;
      a.click();
      URL.revokeObjectURL(url);
      setExportMsg("✓ Newsletter heruntergeladen.");
    } catch (e) { setExportMsg("⚠ Fehler: " + e.message); }
    setExportLoading(false);
  };

  const handleAddProduct = async () => {
    const price = parseFloat(newProduct.price);
    if (!newProduct.name.trim() || isNaN(price) || price <= 0) return;
    try {
      await dbAddProduct({ name: newProduct.name.trim(), price, unit: newProduct.unit.trim(), category_id: newProduct.category_id ?? categories[0]?.id ?? null, special: false });
      await onReloadProducts();
      setShowNewForm(false);
      setNewProduct({ name: "", price: "", unit: "", category_id: null });
    } catch (e) { console.error(e); }
  };

  const totalOrders = orders.length;
  const totalProductsOrdered = new Set(orders.flatMap((o) => Object.keys(o.cart || {}).map(Number))).size;

  return (
    <div style={{ minHeight: "100vh", background: "#f8fafb" }}>
      <header style={{ background: "#14532d", color: "white", padding: "0 1.5rem" }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 22 }}>🌿</span>
            <span style={{ fontWeight: 700, fontSize: 16 }}>Admin · Andreas' Gemüsekisterl</span>
          </div>
          <button onClick={onLogout} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "white", borderRadius: 8, padding: "6px 14px", fontSize: 13, cursor: "pointer" }}>Abmelden</button>
        </div>
      </header>

      <div style={{ maxWidth: 1100, margin: "1.5rem auto", padding: "0 1.5rem" }}>
        <div style={{ display: "flex", gap: 8, marginBottom: "1.5rem" }}>
          {["orders", "products", "kategorien", "upload", "mailchimp"].map((t) => (
            <button key={t} onClick={() => setTab(t)} style={{ background: tab === t ? "#16a34a" : "white", color: tab === t ? "white" : "#4b7c59", border: "1px solid " + (tab === t ? "#16a34a" : "#bbf7d0"), borderRadius: 10, padding: "8px 18px", fontSize: 14, cursor: "pointer", fontWeight: tab === t ? 700 : 400 }}>
              {t === "orders" ? "📋 Bestellungen" : t === "products" ? "🥦 Produktliste" : t === "kategorien" ? "🏷️ Kategorien" : t === "upload" ? "📤 Excel-Import" : "📧 Mailchimp"}
            </button>
          ))}
        </div>

        {tab === "orders" && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: "1.5rem" }}>
              {[
                { label: "Bestellungen", value: totalOrders },
                { label: "Produkte gesamt", value: products.length },
                { label: "Bestellte Artikel", value: totalProductsOrdered },
              ].map((s) => (
                <div key={s.label} style={{ background: "white", borderRadius: 12, padding: "16px 20px", border: "1px solid #e8f5e9", textAlign: "center" }}>
                  <div style={{ fontSize: 28, fontWeight: 800, color: "#16a34a" }}>{s.value}</div>
                  <div style={{ fontSize: 13, color: "#4b7c59" }}>{s.label}</div>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
              <button onClick={onRefresh} style={{ background: "white", border: "1px solid #bbf7d0", color: "#4b7c59", borderRadius: 8, padding: "6px 14px", fontSize: 13, cursor: "pointer" }}>
                🔄 Aktualisieren
              </button>
            </div>
            {orders.length === 0 ? (
              <div style={{ textAlign: "center", padding: "3rem", color: "#4b7c59" }}>Noch keine Bestellungen eingegangen.</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {orders.map((order) => {
                  const cart = order.cart || {};
                  const userName = order.users?.name || "Unbekannt";
                  const total = Object.entries(cart).reduce((s, [id, q]) => {
                    const p = products.find((x) => x.id === parseInt(id));
                    return s + (p ? p.price * q : 0);
                  }, 0);
                  const items = Object.entries(cart)
                    .map(([id, q]) => ({ p: products.find((x) => x.id === parseInt(id)), q }))
                    .filter((x) => x.p);
                  const date = order.updated_at ? new Date(order.updated_at).toLocaleString("de-AT") : "";
                  return (
                    <div key={order.users?.id} style={{ background: "white", borderRadius: 14, padding: "16px 20px", border: "1px solid #e8f5e9" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                        <div>
                          <span style={{ fontWeight: 700, fontSize: 16, color: "#14532d" }}>👤 {userName}</span>
                          {date && <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 10 }}>{date}</span>}
                        </div>
                        <div style={{ fontWeight: 800, color: "#16a34a", fontSize: 16 }}>€ {total.toFixed(2)}</div>
                      </div>
                      {items.length === 0 ? (
                        <span style={{ fontSize: 12, color: "#9ca3af" }}>Leere Bestellung</span>
                      ) : (
                        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                          {items.map(({ p, q }) => (
                            <span key={p.id} style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "4px 10px", fontSize: 12, color: "#166534" }}>
                              {p.name} × {q} {p.unit}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {tab === "products" && (
          <div>
            <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: "1rem", flexWrap: "wrap" }}>
              <input type="text" placeholder="🔍 Produkt suchen…" value={productSearch}
                onChange={(e) => setProductSearch(e.target.value)}
                style={{ ...inputStyle, flex: "1 1 200px", padding: "8px 12px", fontSize: 13 }} />
              <span style={{ color: "#4b7c59", fontSize: 13, whiteSpace: "nowrap" }}>
                {products.filter((p) => p.name.toLowerCase().includes(productSearch.toLowerCase())).length} / {products.length}
              </span>
              <button onClick={() => { setShowNewForm(true); setEditingId(null); }} disabled={showNewForm}
                style={{ background: "#16a34a", color: "white", border: "none", borderRadius: 10, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: showNewForm ? 0.5 : 1 }}>
                + Neues Produkt
              </button>
            </div>

            {showNewForm && (
              <div style={{ background: "#f0fdf4", borderRadius: 12, padding: "16px", border: "2px solid #86efac", marginBottom: "1rem" }}>
                <div style={{ fontWeight: 700, color: "#14532d", marginBottom: 12, fontSize: 14 }}>Neues Produkt hinzufügen</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 8 }}>
                  <input placeholder="Produktname *" value={newProduct.name}
                    onChange={(e) => setNewProduct((p) => ({ ...p, name: e.target.value }))}
                    style={{ ...inputStyle, flex: "1 1 200px", padding: "8px 12px", fontSize: 13 }} />
                  <input type="number" placeholder="Preis (€) *" step="0.1" min="0" value={newProduct.price}
                    onChange={(e) => setNewProduct((p) => ({ ...p, price: e.target.value }))}
                    style={{ ...inputStyle, width: 110, flex: "0 0 110px", padding: "8px 12px", fontSize: 13 }} />
                  <input placeholder="Einheit (z.B. Kilo)" value={newProduct.unit}
                    onChange={(e) => setNewProduct((p) => ({ ...p, unit: e.target.value }))}
                    style={{ ...inputStyle, width: 140, flex: "0 0 140px", padding: "8px 12px", fontSize: 13 }} />
                </div>
                <select value={newProduct.category_id ?? categories[0]?.id ?? ""}
                  onChange={(e) => setNewProduct((p) => ({ ...p, category_id: parseInt(e.target.value) }))}
                  style={{ ...inputStyle, marginBottom: 12, padding: "8px 12px", fontSize: 13 }}>
                  {categories.map((c) => (
                    <option key={c.id} value={c.id}>{categoryIcons[c.name] || "🥦"} {c.name}</option>
                  ))}
                </select>
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={handleAddProduct}
                    style={{ background: "#16a34a", color: "white", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                    Hinzufügen
                  </button>
                  <button onClick={() => { setShowNewForm(false); setNewProduct({ name: "", price: "", unit: "", category_id: null }); }}
                    style={{ background: "white", color: "#4b7c59", border: "1px solid #bbf7d0", borderRadius: 8, padding: "8px 18px", fontSize: 13, cursor: "pointer" }}>
                    Abbrechen
                  </button>
                </div>
              </div>
            )}

            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {products.filter((p) => p.name.toLowerCase().includes(productSearch.toLowerCase())).map((p) => editingId === p.id ? (
                <div key={p.id} style={{ background: "white", borderRadius: 12, padding: "12px 16px", border: "2px solid #86efac", display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                  <input value={editFields.name}
                    onChange={(e) => setEditFields((f) => ({ ...f, name: e.target.value }))}
                    style={{ ...inputStyle, flex: "1 1 180px", padding: "6px 10px", fontSize: 13 }} />
                  <input type="number" step="0.1" min="0" value={editFields.price}
                    onChange={(e) => setEditFields((f) => ({ ...f, price: e.target.value }))}
                    style={{ ...inputStyle, width: 90, flex: "0 0 90px", padding: "6px 10px", fontSize: 13 }} />
                  <input value={editFields.unit}
                    onChange={(e) => setEditFields((f) => ({ ...f, unit: e.target.value }))}
                    style={{ ...inputStyle, width: 120, flex: "0 0 120px", padding: "6px 10px", fontSize: 13 }} />
                  <select value={editFields.category_id ?? ""}
                    onChange={(e) => setEditFields((f) => ({ ...f, category_id: parseInt(e.target.value) }))}
                    style={{ ...inputStyle, flex: "1 1 150px", padding: "6px 10px", fontSize: 13 }}>
                    {categories.map((c) => <option key={c.id} value={c.id}>{categoryIcons[c.name] || "🥦"} {c.name}</option>)}
                  </select>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => handleSaveProduct(p.id)}
                      style={{ background: "#16a34a", color: "white", border: "none", borderRadius: 8, padding: "6px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                      Speichern
                    </button>
                    <button onClick={() => setEditingId(null)}
                      style={{ background: "white", color: "#4b7c59", border: "1px solid #bbf7d0", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>
                      Abbrechen
                    </button>
                  </div>
                </div>
              ) : (
                <div key={p.id} style={{ background: "white", borderRadius: 12, padding: "12px 16px", border: "1px solid #e8f5e9", display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 600, color: "#14532d" }}>{p.name}</div>
                    <div style={{ fontSize: 12, color: "#4b7c59" }}>€ {p.price.toFixed(2)} / {p.unit} · <span style={{ color: "#6ee7b7" }}>{p.category}</span></div>
                  </div>
                  <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                    <button onClick={() => { setEditingId(p.id); setEditFields({ name: p.name, price: p.price, unit: p.unit, category_id: p.category_id }); setShowNewForm(false); }}
                      style={{ background: "white", color: "#4b7c59", border: "1px solid #bbf7d0", borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>
                      Bearbeiten
                    </button>
                    <button onClick={() => handleDeleteProduct(p.id)}
                      style={{ background: "white", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>
                      Löschen
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {tab === "kategorien" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
              <span style={{ fontWeight: 700, color: "#14532d", fontSize: 15 }}>{categories.length} Kategorien</span>
              <button onClick={() => { setShowNewCatForm(true); setEditingCatId(null); }} disabled={showNewCatForm}
                style={{ background: "#16a34a", color: "white", border: "none", borderRadius: 10, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: showNewCatForm ? 0.5 : 1 }}>
                + Neue Kategorie
              </button>
            </div>

            {showNewCatForm && (
              <div style={{ background: "#f0fdf4", borderRadius: 12, padding: "16px", border: "2px solid #86efac", marginBottom: "1rem" }}>
                <div style={{ fontWeight: 700, color: "#14532d", marginBottom: 12, fontSize: 14 }}>Neue Kategorie</div>
                <input placeholder="Name *" value={newCatName}
                  onChange={(e) => setNewCatName(e.target.value)}
                  style={{ ...inputStyle, marginBottom: 10, padding: "8px 12px", fontSize: 13 }} />
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={handleAddCategory}
                    style={{ background: "#16a34a", color: "white", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                    Hinzufügen
                  </button>
                  <button onClick={() => { setShowNewCatForm(false); setNewCatName(""); }}
                    style={{ background: "white", color: "#4b7c59", border: "1px solid #bbf7d0", borderRadius: 8, padding: "8px 18px", fontSize: 13, cursor: "pointer" }}>
                    Abbrechen
                  </button>
                </div>
              </div>
            )}

            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {categories.map((cat) => editingCatId === cat.id ? (
                <div key={cat.id} style={{ background: "white", borderRadius: 12, padding: "12px 16px", border: "2px solid #86efac", display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 20 }}>{categoryIcons[editCatName] || "🥦"}</span>
                  <input value={editCatName}
                    onChange={(e) => setEditCatName(e.target.value)}
                    style={{ ...inputStyle, flex: 1, padding: "6px 10px", fontSize: 13 }} />
                  <button onClick={() => handleRenameCategory(cat)}
                    style={{ background: "#16a34a", color: "white", border: "none", borderRadius: 8, padding: "6px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                    Speichern
                  </button>
                  <button onClick={() => setEditingCatId(null)}
                    style={{ background: "white", color: "#4b7c59", border: "1px solid #bbf7d0", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>
                    Abbrechen
                  </button>
                </div>
              ) : (
                <div key={cat.id} style={{ background: "white", borderRadius: 12, padding: "12px 16px", border: "1px solid #e8f5e9", display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 20 }}>{categoryIcons[cat.name] || "🥦"}</span>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 600, color: "#14532d" }}>{cat.name}</div>
                    <div style={{ fontSize: 11, color: "#4b7c59" }}>{products.filter((p) => p.category === cat.name).length} Produkte</div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => { setEditingCatId(cat.id); setEditCatName(cat.name); setShowNewCatForm(false); }}
                      style={{ background: "white", color: "#4b7c59", border: "1px solid #bbf7d0", borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>
                      Umbenennen
                    </button>
                    <button onClick={() => handleDeleteCategory(cat)}
                      style={{ background: "white", color: "#dc2626", border: "1px solid #fca5a5", borderRadius: 8, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>
                      Löschen
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {tab === "mailchimp" && (
          <div style={{ maxWidth: 480 }}>
            <div style={{ background: "white", borderRadius: 16, padding: "2rem", border: "1px solid #e8f5e9" }}>
              <div style={{ fontSize: 40, marginBottom: "1rem" }}>📧</div>
              <h3 style={{ fontWeight: 700, fontSize: 18, color: "#14532d", marginBottom: 8 }}>Mailchimp-Newsletter exportieren</h3>
              <p style={{ color: "#4b7c59", fontSize: 14, marginBottom: "1.5rem", lineHeight: 1.6 }}>
                Generiert eine HTML-Datei auf Basis der bestehenden Newsletter-Vorlage.
                Die Produktliste wird durch die aktuellen Datenbankprodukte ersetzt.
                Die Datei kann direkt in Mailchimp hochgeladen werden.
              </p>
              <button onClick={handleNewsletterExport} disabled={exportLoading}
                style={{ ...btnPrimary, width: "100%", opacity: exportLoading ? 0.7 : 1 }}>
                {exportLoading ? "Generiere…" : `Newsletter exportieren (${products.length} Produkte)`}
              </button>
              {exportMsg && (
                <div style={{ marginTop: "1rem", padding: "10px 14px", background: exportMsg.startsWith("✓") ? "#f0fdf4" : "#fef9c3", border: "1px solid " + (exportMsg.startsWith("✓") ? "#bbf7d0" : "#fde68a"), borderRadius: 8, fontSize: 13, color: exportMsg.startsWith("✓") ? "#166534" : "#854d0e" }}>
                  {exportMsg}
                </div>
              )}
            </div>
          </div>
        )}

        {tab === "upload" && (
          <div style={{ maxWidth: 480 }}>
            <div style={{ background: "white", borderRadius: 16, padding: "2rem", border: "1px solid #e8f5e9" }}>
              <div style={{ fontSize: 40, marginBottom: "1rem" }}>📤</div>
              <h3 style={{ fontWeight: 700, fontSize: 18, color: "#14532d", marginBottom: 8 }}>Excel-Datei importieren</h3>
              <p style={{ color: "#4b7c59", fontSize: 14, marginBottom: "1.5rem", lineHeight: 1.6 }}>
                Lade die wöchentliche Produktliste hoch (Spalten: B = Name, C = Preis, D = Einheit).
                Bestehende Produkte werden ersetzt.
              </p>
              <button onClick={onUploadClick} disabled={loading}
                style={{ ...btnPrimary, width: "100%", opacity: loading ? 0.7 : 1 }}>
                {loading ? "Importieren…" : "Datei auswählen (.xlsx)"}
              </button>
              {uploadMsg && (
                <div style={{ marginTop: "1rem", padding: "10px 14px", background: uploadMsg.startsWith("✓") ? "#f0fdf4" : "#fef9c3", border: "1px solid " + (uploadMsg.startsWith("✓") ? "#bbf7d0" : "#fde68a"), borderRadius: 8, fontSize: 13, color: uploadMsg.startsWith("✓") ? "#166534" : "#854d0e" }}>
                  {uploadMsg}
                </div>
              )}
            </div>
            {children}
          </div>
        )}
      </div>
    </div>
  );
}

// ── Shared styles & helpers ───────────────────────────────────────────────────

const btnPrimary = {
  background: "#16a34a", color: "white", border: "none",
  borderRadius: 14, padding: "14px 32px", fontSize: 15,
  fontWeight: 700, cursor: "pointer",
};
const btnSecondary = {
  background: "white", color: "#166534",
  border: "2px solid #bbf7d0", borderRadius: 14,
  padding: "14px 32px", fontSize: 15, fontWeight: 600, cursor: "pointer",
};
const backBtn = {
  background: "none", border: "none", color: "#4b7c59",
  cursor: "pointer", fontSize: 14, marginBottom: "1.5rem", display: "block",
};
const inputStyle = {
  border: "1.5px solid #bbf7d0", borderRadius: 10,
  padding: "12px 14px", fontSize: 15, outline: "none",
  color: "#14532d", width: "100%",
};
function ErrorBox({ msg }) {
  return (
    <div style={{ color: "#dc2626", fontSize: 13, padding: "8px 12px", background: "#fef2f2", borderRadius: 8 }}>
      {msg}
    </div>
  );
}
