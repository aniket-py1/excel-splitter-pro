# 🌐 Excel & CSV Splitter Pro — Web Architecture & Monetization Blueprint

This document outlines the architecture, tech stack, monetization model, and implementation roadmap for developing the standalone web version of **Excel & CSV Splitter Pro** (maintained in its own dedicated repository).

---

## 1. Executive Summary

- **Product**: Web-based Excel and CSV dataset splitter.
- **Differentiator**: 100% in-browser client-side processing (Web Workers). Sensitive spreadsheet data **never touches the server**, eliminating cloud server compute/RAM bills ($0 hosting costs) while maximizing user trust.
- **Monetization**: Dual revenue engine:
  1. **Display Advertisements** on the Free Tier (Google AdSense / Carbon Ads).
  2. **Pro Subscription / Lifetime License** ($5/mo or $29 lifetime) unlocking unlimited row processing, multi-file batch splitting, and an ad-free interface.
- **Repository Strategy**: Maintained as an independent repository (e.g., `excel-splitter-web`) decoupled from the PyQt5 desktop client.

---

## 2. Tech Stack Evaluation & Selection

| Layer | Recommended Choice | Rationale |
| :--- | :--- | :--- |
| **Framework** | **Next.js 15+ (App Router)** | Full-stack React with Server-Side Rendering (SSR) for search engine dominance (vital for ranking on *"split csv online"*). Single codebase for UI, API routes, and payment webhooks. |
| **Styling** | **Tailwind CSS & Lucide Icons** | Ultra-clean Obsidian & Slate dark UI matching the desktop suite aesthetic. |
| **Parsing Engine (CSV)** | **PapaParse** (Streamed via Web Worker) | Chunk-by-chunk client-side streaming parser with memory-efficient backpressure. |
| **Parsing Engine (Excel)** | **SheetJS (xlsx) / ExcelJS** | Fast parsing of multi-sheet workbooks in browser RAM. |
| **Concurrency** | **HTML5 Web Workers** | Offloads file reading, partitioning, and zip bundling off the main UI thread to prevent browser tab freezing. |
| **Archiving** | **JSZip** | Client-side compression to package split files into a single `.zip` download. |
| **Hosting** | **Vercel / Cloudflare Pages** | **$0/month** on the free tier with global edge CDN delivery. |
| **Payments** | **Lemon Squeezy / Polar.sh** | Merchant of Record (handles global sales tax, VAT, and invoices automatically). |

---

## 3. Tier Comparison & Feature Limits

| Feature | Free Tier (Ad-Supported) | Pro Tier ($5/mo or $29 Lifetime) |
| :--- | :--- | :--- |
| **Max Rows per File** | 25,000 – 50,000 rows | **Unlimited (1M+ rows)** |
| **Max File Size** | 25 MB | **500 MB+** |
| **Advertisements** | Display Ads (AdSense / EthicalAds) | **100% Ad-Free Clean UI** |
| **Simultaneous Files** | 1 file at a time | **Batch Multi-File Splitting** |
| **Desktop App License** | Not included | **Includes Desktop Pro License Key** |
| **Output Naming** | Standard numeric numbering | Custom naming tokens (`{name}_{date}_{batch}`) |
| **Priority Queue** | Standard Web Worker | Multi-core parallel workers |

---

## 4. Architecture Diagram

```mermaid
graph TD
    User([User Browser]) --> UI[Next.js Modern Web Interface]
    
    subgraph Client-Side Engine (Free Tier - $0 Server Cost)
        UI --> WW[Background Web Worker]
        WW --> CSV[PapaParse Stream]
        WW --> XLS[SheetJS Parser]
        CSV & XLS --> Partitioner[Chunk Partition Engine]
        Partitioner --> Zip[JSZip In-Memory Archiver]
        Zip --> Download([Instant Browser Download])
    end

    subgraph Monetization Layer
        UI --> Ads[Google AdSense / Carbon Banner Ads]
        UI --> Checkout[Lemon Squeezy Checkout]
        Checkout --> Webhook[Next.js Webhook Handler /api/webhook]
        Webhook --> LicenseDB[Supabase / Redis License Registry]
        LicenseDB --> ProKey[Issued Pro License Key]
    end
```

---

## 5. Domain & Ad Network Strategy

1. **Free Subdomains Warning**:
   - Google AdSense **regularly rejects free subdomains** (`.vercel.app`, `.netlify.app`) due to lack of domain ownership proof.
2. **Recommended Action**:
   - Deploy code on Vercel/Cloudflare for **$0**.
   - Connect a cheap custom domain for **$2–$4/year** (e.g. `excelsplitter.xyz`, `datasplitter.online` on Namecheap/Porkbun).
   - This unlocks instant AdSense approval and enterprise credibility for paid customers.

---

## 6. Development Phases for the Web Repo

- [ ] **Phase 1: Project Initialization**: Create standalone Next.js repo with Tailwind CSS and Obsidian theme.
- [ ] **Phase 2: Client-side Engine**: Implement Web Worker pipeline with PapaParse, SheetJS, and JSZip.
- [ ] **Phase 3: Ad Integration**: Configure Google AdSense slots in sidebar and post-split banner.
- [ ] **Phase 4: Billing & Licensing**: Integrate Lemon Squeezy for Pro plan and generate unified desktop + web license keys.
