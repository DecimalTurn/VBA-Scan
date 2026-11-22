import { cp, mkdir } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const site = join(root, "site");
await mkdir(site, { recursive: true });

for (const file of ["index.html", "styles.css", "theme.js", "LICENSE", "CNAME"]) {
  await cp(join(root, file), join(site, file));
}
for (const directory of ["assets", "dist"]) {
  await cp(join(root, directory), join(site, directory), { recursive: true });
}
