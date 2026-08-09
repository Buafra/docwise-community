# DocWise Open Core Strategy

DocWise should have two editions:

1. **DocWise Community** — open source, powerful local-first OCR/RAG app
2. **DocWise Pro / Enterprise** — commercial product with hosted access, teams, support, packaging, managed AI, and business features

## Important decision

Community should be genuinely useful, not crippled. It can include advanced OCR/RAG features when the user brings their own API keys.

Pro should make money from convenience, hosting, team workflows, compliance, support, installers, managed credits, and commercial control.

## Recommended license

For Community:

- **AGPLv3** if you want to prevent SaaS clones from taking the code private
- **Apache 2.0** if you want maximum adoption

Recommended for this project: **AGPLv3**.

Reason: DocWise can become a hosted SaaS. AGPL protects the hosted version from being copied without sharing modifications.

---

## Community vs Pro

| Feature | DocWise Community | DocWise Pro / Enterprise |
|---|---:|---:|
| Local web app | ✅ | ✅ |
| Upload files | ✅ | ✅ |
| Folder indexing | ✅ | ✅ |
| Watched folders | ✅ | ✅ |
| Arabic + English Tesseract OCR | ✅ | ✅ |
| OpenAI Vision OCR | ✅ BYO API key | ✅ Managed credits or BYO key |
| PDF/image/Office support | ✅ | ✅ |
| Manual OCR correction | ✅ | ✅ |
| OCR quality scoring | ✅ | ✅ |
| Smart filing suggestions | ✅ | ✅ |
| Safe copy-to-archive | ✅ | ✅ |
| Advanced hybrid RAG | ✅ | ✅ |
| SQLite FTS5/BM25 | ✅ | ✅ |
| Vector embeddings | ✅ local/BYO OpenAI | ✅ managed or BYO |
| GPT reranking | ✅ BYO API key | ✅ managed credits or BYO |
| Answer verification | ✅ BYO API key | ✅ managed credits or BYO |
| RAG evaluation tests | ✅ | ✅ |
| Arabic query normalization | ✅ | ✅ |
| Structured field extraction | ✅ basic/BYO AI | ✅ advanced schemas + managed AI |
| Astryx UI | ✅ | ✅ |
| Open source code | ✅ | ❌ private/commercial add-ons |
| Sales CRM | ❌ | ✅ |
| License activation | ❌ | ✅ |
| Hosted license server | ❌ | ✅ |
| Stripe/payment integration | ❌ | ✅ |
| Customer portal | ❌ | ✅ |
| Hosted SaaS access | ❌ | ✅ |
| Multi-tenant cloud storage | ❌ | ✅ |
| Team users/roles | ❌ | ✅ |
| Shared team archive | ❌ | ✅ |
| Admin dashboard | ❌ | ✅ |
| Usage metering/quotas | ❌ | ✅ |
| Managed OpenAI/OCR credits | ❌ | ✅ |
| Cloud backup/sync | ❌ | ✅ |
| Desktop installer | Community build/manual | ✅ polished installer |
| Android/mobile packaging | Community build/manual | ✅ packaged/supported |
| Priority support | Community/GitHub | ✅ paid support |
| Custom branding | ❌ | ✅ |
| On-prem deployment support | DIY | ✅ paid setup |
| SLA/commercial contract | ❌ | ✅ |

---

## What Community includes

DocWise Community should include:

- local OCR/RAG archive
- Arabic/English OCR
- folder indexing
- Office/PDF/image support
- OpenAI Vision OCR with user's own key
- advanced hybrid RAG
- GPT reranking with user's own key
- answer verification with user's own key
- smart filing suggestions
- manual OCR correction
- RAG evaluation
- Astryx UI

This makes the open-source version impressive and trustworthy.

---

## What Pro sells

Pro should not sell basic features. It should sell:

### 1. Convenience

- hosted SaaS
- no setup
- managed AI/OCR credits
- polished installers
- updates

### 2. Team workflows

- team accounts
- roles/permissions
- shared archive
- admin dashboard
- audit logs

### 3. Commercial reliability

- license management
- billing
- customer portal
- support
- backups
- monitoring
- SLA

### 4. Enterprise control

- custom branding
- private deployment
- on-prem setup
- compliance support
- custom integrations

---

## Recommended repo structure

Use two repos/folders first:

```txt
C:\Users\YourName\docwise-community
C:\Path\To\docwise-community
```

Later, move to monorepo:

```txt
docwise/
  apps/community
  apps/pro
  apps/commercial-server
  packages/ocr
  packages/rag
  packages/ui
  packages/document-core
```

---

## Recommended pricing after this split

| Plan | Price | Notes |
|---|---:|---|
| Community | Free | open source, BYO keys |
| Pro Local | $19/month or $199 one-time | installer + license + updates |
| Pro SaaS | $39/month | hosted account + managed credits |
| Team | $149/month | multi-user/shared archive |
| Enterprise | custom | on-prem, SLA, integrations |

---

## Best positioning

**DocWise Community**

> Open-source Arabic-first OCR and RAG archive for local documents.

**DocWise Pro**

> Managed, supported and team-ready DocWise with hosted access, licenses, billing, cloud features and enterprise support.
