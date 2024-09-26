import * as esbuild from "esbuild"

await esbuild.build({
  entryPoints: ["src/index.ts"],
  bundle: true,
  globalName: "TheHopeCalendar",
  outfile: "dist.js",
})
