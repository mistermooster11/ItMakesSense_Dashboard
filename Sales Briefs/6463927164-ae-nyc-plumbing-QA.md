# QA Report — A&E NYC Plumbing
## Template: Pipe Monkeys   Client: AE-NYC-Plumbing
## Run: 2026-05-26

---

### 🗂 Page Inventory
| Page | Template | Client | Status |
|------|----------|--------|--------|
| / (homepage) | ✅ | ✅ | ✅ |
| /blog | ✅ | ✅ | ✅ |
| /contact-us | ✅ | ✅ | ✅ |
| /craft-catalog | ✅ | ✅ | ✅ |
| /craft-catalog/[slug] | ✅ | ✅ | ✅ |
| /explore/[slug] | ✅ | ✅ | ✅ |
| /gallery | ✅ | ✅ | ✅ |
| /general-faqs | ✅ | ✅ | ✅ |
| /programs-crafts/programs | ✅ | ✅ | ✅ |
| /service-areas | ✅ | ✅ | ✅ |
| /about-us | ➕ client-only | ✅ | ✅ added |
| /faq | ➕ client-only | ✅ | ✅ added |
| /drain-repair | ➕ client-only | ✅ | ✅ added |
| /emergency-plumbing | ➕ client-only | ✅ | ✅ added |
| /leak-detection | ➕ client-only | ✅ | ✅ added |
| /services-page | ➕ client-only | ✅ | ✅ added |
| /toilet-repair | ➕ client-only | ✅ | ✅ added |
| /water-heater | ➕ client-only | ✅ | ✅ added |
| /privacy-policy | ➕ client-only | ✅ | ✅ added |
| /blog/blog-unclogme | ➕ client-only | ✅ | ✅ added |
| /craft-catalog/pipeline-corrosion-control | template-only | ❌ removed | ✅ expected (template-specific page) |

Missing pages: None — all template pages present. Client added 10 service/utility pages beyond template baseline.

---

### 🏗 Structure — Page by Page
**/ (homepage)** ✅
- FAQSection, TestimonialsSection, HeroSection, ServicesSection, CTAFormSection all present
- No structural drift detected

**All other pages** ✅
- Standard page structure confirmed via import audit

---

### 🏷 Metadata
- layout.tsx title template: ✅ `"%s | A&E NYC Plumbing"` — properly customized
- Default "Create Next App" strings in user-facing files: ✅ none found
- ⚠️ README.md still contains default Next.js boilerplate — not user-facing, but worth replacing if the repo is shared

---

### 🔒 Brand Leak — Pipe Monkeys References
- ❌ **`data/channel/pipemonkeys.tsx`** — file exists with:
  - Variable name: `const pipemonkeys`
  - Slug: `"pipemonkeys"`
  - User-facing copy: "Pipe Monkeys is a family-owned drain and sewer company..." (3+ instances)
- Status: File is **NOT imported** by any page in `app/` — orphaned dead file. Not currently rendered.
- Fix: **Delete** `data/channel/pipemonkeys.tsx` before launch. An orphaned template data file with brand copy is a latent liability.

---

### 📞 Phone Number Sweep
- Only phone number found in `app/`: `(646)` — A&E's own area code prefix ✅
- No foreign phone numbers detected

---

### 🔗 Slug Consistency
- crafts.ts slugs (16): `dishwasher-repair`, `drain-repair`, `faucets-and-sinks`, `garbage-disposal`, `gas-line-services`, `kitchen-and-bath-plumbing`, `leak-detection`, `piping-and-repiping`, `plumbing-fixtures`, `residential-plumbing`, `shower-and-tub`, `toilet-repair`, `water-filtration`, `water-heater-installation`, `water-line-services`, `water-pressure-repair`
- service-pages.tsx keys: **Exact match** ✅
- Mismatches: None

---

### 📁 Nested Template Folder
- No nested template detected ✅

---

### ✏️ TODO Items Remaining
| File | Count |
|------|-------|
| app/blog/page.tsx | 2 |
| app/gallery/page.tsx | 4 |
| components/custom/Footer.tsx | 2 |

⚠️ 8 TODOs remaining — expected until real blog content, gallery images, and footer details are provided. Flag these as pending before launch.

---

### 🤖 AI Visibility (AEO Checks)
| Check | Result |
|-------|--------|
| JSON-LD schema (LocalBusiness/Service) in layout.tsx | ✅ found — `application/ld+json` with `localBusinessSchema` in layout.tsx |
| FAQ section or component | ✅ found — `FAQSection` on homepage, dedicated `/faq` page, `/general-faqs` page |
| E-E-A-T signals (license, owner, years, insured) | ✅ present in 8 files (layout.tsx, about-us, emergency-plumbing, general-faqs, service-areas, water-heater) |

> ℹ️ All three AEO checks pass. This build is structured to be discoverable by AI answer engines. Schema tells AI tools what the business is; FAQ pages give them content to cite; E-E-A-T signals across 8 files establish authority and trustworthiness.

---

### 📋 Summary
| Check | Result |
|-------|--------|
| Page inventory | ✅ All template pages present + 10 additional service pages |
| Page structure | ✅ No structural drift |
| Metadata | ✅ Properly customized |
| Brand leaks | ❌ `data/channel/pipemonkeys.tsx` — orphaned, not rendered, but must be deleted before launch |
| Phone sweep | ✅ No foreign numbers |
| Slug consistency | ✅ Perfect match (16/16) |
| Nested folder | ✅ None |
| TODOs | ⚠️ 8 items across 3 files (blog, gallery, footer — expected) |
| AI visibility (AEO) | ✅ All 3 pass — schema, FAQ, E-E-A-T all present |

**Overall: FAIL — 1 fix required before sending.**

Fix `data/channel/pipemonkeys.tsx`:
```bash
rm "Clients/AE-NYC-Plumbing/data/channel/pipemonkeys.tsx"
```
After that: PASS (TODOs are expected pending asset delivery).
