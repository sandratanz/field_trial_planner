# Field Trial Planner (GRDC-West)

A Shiny app that collects the information a grower/researcher needs to
provide before a trial can be designed, following AAGI's Analytics
Collaboration Plan (ACP). 

**Researchers don't need R installed.** This app is deployed with
[shinylive](https://posit-dev.github.io/r-shinylive/), which compiles the
app (R and all) to WebAssembly and runs it entirely in the visitor's
browser. There's no server behind it — it's just a website.

## One-time setup (you, not researchers)

1. Create a new GitHub repository and add these files to it (`app.R`, the
   `.github/workflows/deploy.yaml` file, this README, `.gitignore`) —
   keeping the folder structure as-is.
2. Push to the `main` branch.
3. In the repo, go to **Settings → Pages → Build and deployment → Source**
   and choose **GitHub Actions**.
4. That's it. The push you just made (or the next one) triggers the
   `Deploy Shinylive app to GitHub Pages` workflow under the **Actions**
   tab. Give it a few minutes the first time — it installs R and the
   app's packages, compiles everything to WebAssembly, and publishes the
   result. You can watch its progress under **Actions**.
5. Once it finishes, your app is live at:

   ```
   https://<your-username>.github.io/<your-repo-name>/
   ```

   (GitHub also shows this URL under **Settings → Pages** once it's live.)

From then on, every push to `main` automatically rebuilds and redeploys —
you never need to run R locally, even to deploy.

## Sharing it with researchers

Just send them the `github.io` link above. They open it in any modern
browser (Chrome, Edge, Firefox, Safari) — nothing to install. The first
visit takes a few seconds longer while the browser downloads the R
runtime and packages; after that it's cached and loads quickly.

## What the app covers

The app mirrors the ACP structure:

| Tab | ACP Section | What's collected |
|---|---|---|
| 1. Project Overview | Section 1 | Project title, ID, PI, contact, overall objective |
| 2.1 Trials | Section 2.1 | One row per trial: name, aim, support requested |
| 2.2 Factors & Levels | Section 2.2 | Treatment factors per trial, and a description of every level |
| 2.4 Responses | Section 2.4 | Response variables, units, data type, measurement method, sampling, timing |
| 2.5 Questions | Section 2.5 | Research questions built from response + effect type + factor(s), with an auto-generated question sentence |
| 2.6 Implementation | Section 2.6 | Every site x year implementation, replicates, layout, deviations |
| 2.7 Summary | Section 2.7 | Auto-calculated: implementations, designs, analyses, and complexity flags per trial |
| Roles & Sign-off | Section 3–4 | Static responsibilities text (from the template) plus a sign-off block |
| Save, Load & Download | — | Save all progress to a `.rds` file, reload it later, and export everything as one Excel workbook for AAGI |

Dropdowns for layout type, data type, effect type, etc. are taken directly
from the "Lists" sheet in `ACP_interactive_RM.xlsx`.

## Updating the app

Edit `app.R` and push to `main` — the workflow rebuilds and redeploys
automatically. No separate build step to run yourself.

## Running it locally (optional, only if you have R)

This is only useful for development/testing before you push — researchers
never need to do this.

```r
install.packages(c("shiny", "DT", "openxlsx"))
shiny::runApp("app.R")
```

## Notes / things you may want to tweak

- **Section 2.3 (Replicates)** in the spreadsheet is guidance text rather
  than a data-entry table, so it isn't a separate tab — the number of
  replicates is instead captured per implementation in Section 2.6.
- **Section 2.7 (Power and Precision)** in the spreadsheet is marked as an
  optional future addition, not built here.
- **Save/Load progress** works the same way whether the app is running
  locally or via GitHub Pages — it downloads/reads a `.rds` file through
  the browser, no server storage involved.
- If the in-browser (WebAssembly) version ever feels too slow, or you add
  a package that isn't available for WebAssembly, the fallback is to host
  the same `app.R` on [shinyapps.io](https://www.shinyapps.io/) instead —
  a real R server, still no install required for researchers, just a
  different URL than GitHub Pages.
