#!/usr/bin/env node
// Собирает app/index.html из app/template.html и данных в data/.
// Запуск: node scripts/build.mjs
import { readFileSync, writeFileSync, readdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const read = (p) => JSON.parse(readFileSync(join(root, p), "utf8"));

const profile = read("data/profile.json");
const foods = read("data/foods.json").foods;
const dayFiles = readdirSync(join(root, "data/days")).filter((f) => f.endsWith(".json")).sort();
const days = dayFiles.map((f) => read(join("data/days", f)));

// Проверка: все id из дневника есть в базе продуктов.
const known = new Set(foods.map((f) => f.id));
const missing = new Set();
for (const day of days)
  for (const meal of day.meals)
    for (const item of meal.items) if (!known.has(item.id)) missing.add(`${day.date}: ${item.id}`);
if (missing.size) {
  console.error("Продукты отсутствуют в data/foods.json:\n  " + [...missing].join("\n  "));
  process.exit(1);
}

const payload = JSON.stringify({ profile, foods, days }).replace(/</g, "\\u003c");
const template = readFileSync(join(root, "app/template.html"), "utf8");
if (!template.includes("__DATA__")) {
  console.error("В app/template.html нет плейсхолдера __DATA__");
  process.exit(1);
}
writeFileSync(join(root, "app/index.html"), template.replace("__DATA__", payload));

console.log(`Готово: app/index.html — ${foods.length} продуктов, ${days.length} дней (${days[0].date} … ${days.at(-1).date}).`);
