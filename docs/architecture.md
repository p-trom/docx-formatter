# Założenia architektury SaaS: Smart Document Formatter

## 1. Cele i zakres produktu

### 1.1 Problem użytkownika
Użytkownik posiada:
- **Document A (Template)** — dokument wzorcowy z brandingiem, stylami, układem, logiką formatowania
- **Document B (Content)** — dokument z treścią, często pozbawiony profesjonalnego formatowania

Chce otrzymać **Document C (Output)** — treść z Document B w formacie i stylach Document A.

### 1.2 Przykłady użycia
| Template (A) | Treść (B) | Wynik (C) |
|-------------|-----------|-----------|
| Szablon oferty handlowej (logo, kolory, nagłówki) | Notatki z rozmowy z klientem | Profesjonalna oferta PDF/DOCX |
| Szablon CV (sekcje, style, układ) | Tekst CV z LinkedIn/plain text | CV zgodne z brandingiem firmy |
| Template raportu (TOC, nagłówki, stopki) | Surowe dane z analizy | Raport gotowy do prezentacji |
| Instrukcja formatowania (opis słowny) | Artykuł naukowy | Sformatowany artykuł |

---

## 2. Wymagania funkcjonalne

### 2.1 Poziom MVP (Minimal Viable Product)
- [ ] Upload 2 plików DOCX przez web UI / API
- [ ] Ekstrakcja stylów paragrafów i czcionek z template'u
- [ ] Aplikacja stylów do treści na podstawie nazw stylów (Heading 1 → Heading 1)
- [ ] Zachowanie struktury sekcji (nagłówki, listy, tabele)
- [ ] Pobranie wynikowego pliku DOCX
- [ ] Podgląd wyniku w przeglądarce (render HTML/PNG)

### 2.2 Poziom zaawansowany (v2)
- [ ] Inteligentne mapowanie stylów gdy nazwy się nie zgadzają (AI matching)
- [ ] Obsługa placeholderów ({{nazwa_firmy}}, {{data}})
- [ ] Ekstrakcja i aplikacja stylów tabel
- [ ] Kopiowanie nagłówków/stopek i ustawień strony z template'u
- [ ] Obsługa obrazów (logo w nagłówku, tła)
- [ ] Batch processing (wiele dokumentów na raz)

### 2.3 Poziom enterprise (v3)
- [ ] Template marketplace — biblioteka gotowych szablonów
- [ ] Template designer — WYSIWYG edytor szablonów
- [ ] Reguły biznesowe per template (walidacja, wymagane sekcje)
- [ ] Integracje: Google Docs, SharePoint, Notion, Dropbox
- [ ] White-label API dla innych SaaSów

---

## 3. Architektura systemu

### 3.1 Diagram komponentów (high-level)

```
┌─────────────────────────────────────────────────────────────────┐
│                        CLIENT LAYER                              │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────────┐  │
│  │   Web App    │  │   REST API   │  │   Plugin/Widget      │  │
│  │  (React/Vue) │  │   (B2B/API)  │  │   (Google Docs etc)  │  │
│  └──────┬───────┘  └──────┬───────┘  └──────────┬───────────┘  │
└─────────┼─────────────────┼─────────────────────┼──────────────┘
          │                 │                     │
          └─────────────────┼─────────────────────┘
                            │ HTTPS/WSS
┌───────────────────────────▼───────────────────────────────────┐
│                      API GATEWAY                                 │
│         (Kong / AWS API Gateway / Cloudflare)                   │
│         • Rate limiting • Auth • SSL termination                │
└───────────────────────────┬───────────────────────────────────┘
                            │
┌───────────────────────────▼───────────────────────────────────┐
│                   APPLICATION LAYER                              │
│                                                                  │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐ │
│  │ Document Upload │  │ Job Orchestrator│  │   AI Engine     │ │
│  │    Service      │  │   (Celery/RQ)   │  │  (OpenRouter)   │ │
│  └────────┬────────┘  └────────┬────────┘  └────────┬────────┘ │
│           │                    │                    │          │
│  ┌────────▼────────┐  ┌────────▼────────┐  ┌────────▼────────┐ │
│  │  Template       │  │  Style Matcher  │  │  Content        │ │
│  │  Analyzer       │  │  Engine         │  │  Optimizer      │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘ │
│                                                                  │
└───────────────────────────┬───────────────────────────────────┘
                            │
┌───────────────────────────▼───────────────────────────────────┐
│                  DOCUMENT ENGINE LAYER                           │
│                                                                  │
│  ┌──────────────────────┐  ┌──────────────────────────────┐   │
│  │  DOCX Processor      │  │  Style Mapping Engine        │   │
│  │  (python-docx +      │  │  • Rule-based mapping        │   │
│  │   custom XML ops)    │  │  • LLM-based matching        │   │
│  └──────────────────────┘  │  • Hybrid approach           │   │
│                            └──────────────────────────────┘   │
│  ┌──────────────────────┐  ┌──────────────────────────────┐   │
│  │  Document Assembler  │  │  Placeholder Resolver        │   │
│  │  (generates output)  │  │  • {{vars}} • Conditions     │   │
│  └──────────────────────┘  └──────────────────────────────┘   │
│                                                                  │
└───────────────────────────┬───────────────────────────────────┘
                            │
┌───────────────────────────▼───────────────────────────────────┐
│                     DATA LAYER                                   │
│                                                                  │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐        │
│  │   PostgreSQL │  │    Redis     │  │ Object Store │        │
│  │  (metadata,  │  │  (queues,    │  │  (S3/MinIO)  │        │
│  │   jobs, users│  │   sessions)  │  │  (files,     │        │
│  └──────────────┘  └──────────────┘  │   previews)  │        │
│                                      └──────────────┘        │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## 4. Szczegóły techniczne komponentów

### 4.1 Parsing Layer — Dokument wejściowy

**Problem techniczny:** DOCX to ZIP z XML (Office Open XML). Struktura jest złożona:
- `document.xml` — treść
- `styles.xml` — definicje stylów
- `numbering.xml` — definicje list numerowanych
- `header*.xml`, `footer*.xml` — nagłówki/stopki
- `theme*.xml` — kolory, czcionki motywu
- `settings.xml` — marginesy, orientacja
- `_rels/` — relacje między częściami

**Rozwiązanie:**
```python
# Core libraries
python-docx==1.1.2       # Podstawowa manipulacja DOCX
lxml==5.3.0              # Niskopoziomowy dostęp do XML
zipfile                  # Wbudowane — DOCX to ZIP

# Custom parser dla zaawansowanych przypadków
class DOCXStructureExtractor:
    def extract_template_profile(self, docx_path) -> TemplateProfile:
        """
        Zwraca:
        - paragraph_styles: [{name, font, size, color, spacing, alignment, based_on}]
        - character_styles: [{name, font, size, bold, italic, color}]
        - table_styles: [{name, borders, shading}]
        - document_defaults: {margins, page_size, orientation}
        - theme: {colors[], fonts[]}
        - headers_footers: [{type, content, style_refs}]
        - sections: [{break_type, cols, margins}]
        """
        
    def extract_content_profile(self, docx_path) -> ContentProfile:
        """
        Zwraca:
        - paragraphs: [{text, style_name, level, indentation, is_list}]
        - tables: [{rows, cells, styles}]
        - images: [{embed_id, size, position}]
        - structure_tree: TreeNode[Section]
        """
```

### 4.2 Template Analyzer — Inteligentna analiza template'u

Kluczowe wyzwanie: template może być:
1. **Pure style template** — tylko definicje stylów, pusta treść
2. **Example document** — sformatowany dokument z "lorem ipsum" jako wzór
3. **Instruction-based** — dokument zawierający tekstowe instrukcje formatowania

**Algorytm analizy:**
```python
class TemplateAnalyzer:
    def analyze(self, docx_path) -> TemplateProfile:
        profile = self.extractor.extract_template_profile(docx_path)
        
        # 1. Detekcja typu template'u
        if self.is_style_only(profile):
            profile.type = TemplateType.STYLE_DEFINITIONS
        elif self.has_example_content(profile):
            profile.type = TemplateType.EXAMPLE_DOCUMENT
            profile.semantic_mapping = self.infer_semantic_roles(profile)
        elif self.has_instructions(profile):
            profile.type = TemplateType.INSTRUCTION_BASED
            profile.instructions = self.parse_instructions(profile)
        
        # 2. Budowa mapy stylów semantycznych
        profile.semantic_styles = self.build_semantic_map(profile)
        
        # 3. Detekcja placeholderów
        profile.placeholders = self.detect_placeholders(profile)
        
        return profile
    
    def infer_semantic_roles(self, profile) -> Dict[str, SemanticRole]:
        """
        Przypisuje role semantyczne stylom na podstawie treści przykładowej:
        - "Nagłówek oferty" → role=TITLE
        - "Data: 01.01.2024" → role=DATE_FIELD
        - "Wysokość: 10000 PLN" → role=AMOUNT_FIELD
        """
        # Rule-based + LLM hybrid
```

### 4.3 Style Matching Engine — Mapowanie stylów

**Scenariusze do obsługi:**

| Scenariusz | Przykład | Strategia |
|-----------|----------|-----------|
| Exact match | "Heading 1" → "Heading 1" | Bezpośrednie mapowanie |
| Name similarity | "Nagłówek 1" → "Heading 1" | Fuzzy string matching (Levenshtein + embeddings) |
| Semantic match | Treść wyróżniona → "Quote" | LLM classification |
| Structure-based | Poziom 1 outline → "Heading 1" | Analiza hierarchii (outline level) |
| Content-based | Duży tekst + bold → "Title" | Heurystyki + ML |

**Architektura silnika:**
```python
class StyleMatchingEngine:
    def __init__(self):
        self.rules = RuleBasedMatcher()      # Szybkie, deterministyczne
        self.similarity = EmbeddingMatcher() # Semantic similarity
        self.llm = LLMStyleMatcher()         # Fallback dla edge cases
    
    def match(self, source_style, template_styles) -> MatchResult:
        # 1. Sprawdź exact match
        if match := self.rules.exact_match(source_style, template_styles):
            return match
        
        # 2. Sprawdź fuzzy name match
        if match := self.rules.fuzzy_name_match(source_style, template_styles):
            return match
        
        # 3. Sprawdź outline level match
        if match := self.rules.outline_level_match(source_style, template_styles):
            return match
        
        # 4. Semantic embedding match
        if match := self.similarity.find_best(source_style, template_styles):
            if match.confidence > 0.85:
                return match
        
        # 5. LLM-based matching (dla trudnych przypadków)
        return self.llm.match(source_style, template_styles)
```

**Prompt dla LLM matching:**
```
Given a source paragraph with style "{source_style.name}" and text "{source_text}", 
and template styles: {template_styles_descriptions},
which template style best matches? Return JSON: {"style_name": "...", "confidence": 0.95, "reason": "..."}
```

### 4.4 Document Assembler — Generowanie wyniku

```python
class DocumentAssembler:
    def assemble(self, template_profile, content_profile, style_mapping) -> bytes:
        """
        1. Tworzy nowy dokument oparty na template (kopie definicji stylów)
        2. Iteruje po paragrafach z content
        3. Dla każdego paragrafu: znajduje mapped style z template
        4. Aplikuje style (paragrafowy + znakowy)
        5. Zachowuje strukturę (listy, tabele, sekcje)
        6. Kopiuje nagłówki/stopki z template
        7. Aplikuje marginesy i ustawienia strony
        """
        
    def apply_template_settings(self, output_doc, template_profile):
        """Marginesy, rozmiar strony, orientacja, kolumny"""
        
    def copy_headers_footers(self, output_doc, template_profile):
        """Nagłówki i stopki z template'u (z placeholder replacement)"""
        
    def build_numbering(self, output_doc, template_profile, content_profile):
        """Odbudowuje listy numerowane zgodnie z definicją z template"""
```

---

## 5. Przepływ danych (Data Flow)

```
Użytkownik uploaduje:
├── template.docx ──┐
│                   │     ┌──────────────────┐
│                   ├────►│ Document Upload  │
│                   │     │ Service          │
├── content.docx ───┘     └────────┬─────────┘
                                    │
                          ┌─────────▼─────────┐
                          │   File Storage    │
                          │   (S3/MinIO)      │
                          └─────────┬─────────┘
                                    │
                          ┌─────────▼─────────┐
                          │  Job Orchestrator │
                          │  (publikuje task) │
                          └─────────┬─────────┘
                                    │
                    ┌───────────────┼───────────────┐
                    │               │               │
            ┌───────▼──────┐ ┌─────▼──────┐ ┌─────▼──────┐
            │ Worker 1     │ │ Worker 2   │ │ Worker N   │
            │ (Celery)     │ │ (Celery)   │ │ (Celery)   │
            └───────┬──────┘ └─────┬──────┘ └─────┬──────┘
                    │              │              │
                    └──────────────┼──────────────┘
                                   │
                    ┌──────────────▼──────────────┐
                    │     Processing Pipeline     │
                    │                             │
                    │  1. Parse Template          │
                    │  2. Parse Content           │
                    │  3. Match Styles            │
                    │  4. Resolve Placeholders    │
                    │  5. Assemble Output         │
                    │  6. Generate Preview        │
                    │  7. Store Result            │
                    └──────────────┬──────────────┘
                                   │
                    ┌──────────────▼──────────────┐
                    │  PostgreSQL (job status)    │
                    │  S3 (result.docx, preview)  │
                    └─────────────────────────────┘
                                   │
                    ┌──────────────▼──────────────┐
                    │  WebSocket / SSE / Polling  │
                    │  (notify user: done)        │
                    └─────────────────────────────┘
```

---

## 6. Baza danych — schemat (PostgreSQL)

```sql
-- Użytkownicy i organizacje (multi-tenant)
CREATE TABLE organizations (
    id UUID PRIMARY KEY,
    name TEXT NOT NULL,
    plan TEXT DEFAULT 'free', -- free, pro, enterprise
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE users (
    id UUID PRIMARY KEY,
    email TEXT UNIQUE NOT NULL,
    org_id UUID REFERENCES organizations(id),
    role TEXT DEFAULT 'member', -- admin, member
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Szablony (template library)
CREATE TABLE templates (
    id UUID PRIMARY KEY,
    org_id UUID REFERENCES organizations(id),
    name TEXT NOT NULL,
    description TEXT,
    file_path TEXT NOT NULL, -- S3 key
    file_size INT,
    template_type TEXT, -- style_only, example, instruction_based
    profile JSONB, -- cache parsed template profile
    is_public BOOLEAN DEFAULT false,
    usage_count INT DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Zadania przetwarzania
CREATE TABLE jobs (
    id UUID PRIMARY KEY,
    user_id UUID REFERENCES users(id),
    org_id UUID REFERENCES organizations(id),
    template_id UUID REFERENCES templates(id),
    status TEXT DEFAULT 'pending', -- pending, parsing, matching, assembling, completed, failed
    
    -- Pliki wejściowe
    template_file_path TEXT,
    content_file_path TEXT,
    
    -- Wynik
    output_file_path TEXT,
    output_file_size INT,
    preview_url TEXT, -- HTML/PNG preview
    
    -- Metryki
    style_matches JSONB, -- log decyzji matchingowych
    placeholders_found TEXT[],
    placeholders_resolved JSONB,
    
    -- Czasy
    queued_at TIMESTAMPTZ DEFAULT NOW(),
    started_at TIMESTAMPTZ,
    completed_at TIMESTAMPTZ,
    duration_ms INT,
    
    -- Błędy
    error_message TEXT,
    error_details JSONB,
    
    -- Rate limiting
    retry_count INT DEFAULT 0
);

-- Wersjonowanie stylów (dla template marketplace)
CREATE TABLE template_versions (
    id UUID PRIMARY KEY,
    template_id UUID REFERENCES templates(id),
    version TEXT NOT NULL,
    file_path TEXT NOT NULL,
    change_notes TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);
```

---

## 7. Stack technologiczny

### 7.1 Backend
| Komponent | Technologia | Uzasadnienie |
|-----------|-------------|--------------|
| API Framework | FastAPI (Python) | Szybki, auto-docs, async native |
| Task Queue | Celery + Redis | Sprawdzony, monitoring, retry logic |
| Dokumenty | python-docx + lxml | Standard dla DOCX, pełna kontrola XML |
| AI/LLM | OpenRouter API | Dostęp do Claude, GPT-4, Gemini — fallback |
| Baza danych | PostgreSQL 16 | JSONB dla elastycznych profili |
| Storage | S3-compatible (MinIO/AWS) | Obiekty, presigned URLs |
| Preview | LibreOffice headless | Konwersja DOCX → PDF → PNG preview |
| Auth | Clerk/Auth0 + JWT | OAuth2, SSO, organizacje |

### 7.2 Frontend
| Komponent | Technologia | Uzasadnienie |
|-----------|-------------|--------------|
| Framework | React 18 + TypeScript | Ekosystem, komponenty UI |
| Styling | Tailwind CSS | Szybki development |
| Upload | react-dropzone + tus | Resumable uploads, progres |
| Preview | PDF.js + Mammoth.js | Preview DOCX w przeglądarce |
| State | Zustand | Lekki, nie Redux |

### 7.3 Infrastructure
| Komponent | Technologia |
|-----------|-------------|
| Hosting | Fly.io / AWS / Hetzner |
| Konteneryzacja | Docker + docker-compose (dev), k8s (prod) |
| CI/CD | GitHub Actions |
| Monitoring | Sentry (errors) + Prometheus/Grafana (metrics) |

---

## 8. Bezpieczeństwo i compliance

### 8.1 Bezpieczeństwo dokumentów
- **Isolacja tenantów** — każda organizacja ma osobny prefix w S3
- **Szyfrowanie rest** — S3 SSE-S3 lub SSE-KMS
- **Szyfrowanie in transit** — TLS 1.3
- **Autodestrukcja** — pliki tymczasowe usuwane po 7 dniach (configurable)
- **Watermarking preview** — podglądy z watermarkiem dla planu free

### 8.2 Rate Limiting
| Plan | Jobs/hour | Max file size | Max pages |
|------|-----------|---------------|-----------|
| Free | 10 | 5 MB | 50 |
| Pro | 100 | 25 MB | 200 |
| Enterprise | ∞ | 100 MB | ∞ |

---

## 9. ROI i model biznesowy

### 9.1 Pricing tiers
- **Free** — 10 dokumentów/miesiąc, watermark, podstawowe template'y
- **Pro ($19/user/mo)** — bez limitów watermark, custom templates, API access
- **Enterprise ($49/user/mo)** — SSO, audit log, custom integrations, SLA

### 9.2 Koszty infrastruktury (szacunkowe, per 1000 dokumentów)
| Pozycja | Koszt |
|---------|-------|
| LLM calls (OpenRouter) | $2-5 |
| Compute (Celery workers) | $1-2 |
| Storage (S3) | $0.10 |
| Bandwidth | $0.50 |
| **Razem** | **$4-8** |

Przy cenie Pro: 1000 dokumentów × ~$0.05 = **$50 revenue → ~$42 margin**

---

## 10. Roadmap

| Faza | Czas | Funkcjonalność |
|------|------|---------------|
| **MVP** | 4 tyg. | Upload, basic style matching, download, preview |
| **v1.1** | +2 tyg. | AI matching, placeholders, template library |
| **v1.2** | +2 tyg. | API publiczne, webhooks, batch processing |
| **v2.0** | +4 tyg. | Template designer, integracje (GDocs, Notion) |
| **v3.0** | +6 tyg. | White-label, marketplace, enterprise features |

---

## 11. Ryzyka i ich mitigacja

| Ryzyko | Prawdopodobieństwo | Mitigacja |
|--------|-------------------|-----------|
| Complex DOCX breaks parser | Wysokie | Fallback do LibreOffice, graceful degradation |
| LLM costs spiral | Średnie | Cachowanie wyników matchingowych, rule-based first |
| Format fidelity issues | Wysokie | Preview wymagany przed download, rollback |
| Multi-tenancy leaks | Niskie | Prefix isolation, testy integracyjne |
| Zależność od python-docx | Średnie | Abstract wrapper, możliwość zamiany na .NET/interop |

---

*Dokument przygotowany jako założenia architektoniczne MVP produktu SaaS do inteligentnego formatowania dokumentów DOCX.*
