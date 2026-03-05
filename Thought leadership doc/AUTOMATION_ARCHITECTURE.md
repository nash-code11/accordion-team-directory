# Accordion Insights Automation: Complete Architecture Document

## Table of Contents

1. [What This Document Covers](#1-what-this-document-covers)
2. [Architecture Overview](#2-architecture-overview)
3. [Component Breakdown](#3-component-breakdown)
4. [Data Flow with Example Payloads](#4-data-flow-with-example-payloads)
5. [Column Mapping Logic (All 17 Columns)](#5-column-mapping-logic-all-17-columns)
6. [Implementation Option A: Power Automate + Azure Function](#6-implementation-option-a-power-automate--azure-function)
7. [Implementation Option B: Fully Python-Based Azure Functions](#7-implementation-option-b-fully-python-based-azure-functions)
8. [Error Handling](#8-error-handling)
9. [AI Prompt Engineering](#9-ai-prompt-engineering)
10. [Deployment Checklist](#10-deployment-checklist)

---

## 1. What This Document Covers

This document describes an end-to-end automation system that:

1. **Monitors** the Accordion website (`accordion.com/insights`) for newly published articles, white papers, event recaps, multimedia, and press releases.
2. **Extracts** structured data from each new piece of content (title, author, date, body text, PDFs, external publication links).
3. **Sends the extracted text to an AI model** (Azure OpenAI or OpenAI) to generate a summary, Q&A pairs, keyword tags, audience classification, industry tags, geography tags, solutions/value creation levers, technology/AI relevance, and a client-facing business development email draft.
4. **Writes a new row** into a SharePoint-hosted Excel spreadsheet with all 17 required columns populated.

### Who This Is For

This document is written for someone who may have no prior experience with automation, APIs, or cloud functions. Every concept is explained from the ground up. If you already know what a REST API is, feel free to skip ahead to the architecture diagram.

### Key Concepts (Glossary)

| Term | Plain-English Meaning |
|---|---|
| **REST API** | A URL you can visit (or a program can call) to get structured data back, instead of a visual web page. Think of it like a vending machine: you put in a specific request, and you get back a specific, predictable response. |
| **JSON** | The format that APIs return data in. It looks like nested key-value pairs: `{"name": "John", "age": 30}`. |
| **Azure Function** | A small piece of code that lives in Microsoft's cloud. It runs only when triggered (on a schedule, or by an event) and you only pay for the seconds it runs. |
| **Power Automate** | Microsoft's no-code/low-code automation tool. You build "flows" by connecting triggers and actions like building blocks. |
| **SharePoint** | Microsoft's cloud-based document and collaboration platform. Your Excel spreadsheet lives here. |
| **WordPress REST API** | Accordion's website runs on WordPress, which automatically exposes a REST API. This is how we read article data without manually visiting the website. |
| **Taxonomy** | WordPress organizes content using "taxonomies" -- categories and tags. Each taxonomy term has a numeric ID. For example, `knowledge_type: 44` means "Articles." |
| **Azure OpenAI** | Microsoft's hosted version of OpenAI's GPT models. It is the same AI, but runs inside Azure's infrastructure, which is important for enterprise security and compliance. |

---

## 2. Architecture Overview

### High-Level Flow

```
+---------------------+
|   TRIGGER            |
|   (Every 6 hours)    |
|   Power Automate     |
|   Recurrence         |
|   OR Azure Timer     |
+----------+----------+
           |
           v
+----------+----------+
|   DATA FETCHER       |
|   Call WP REST API   |
|   GET /knowledge     |
|   ?per_page=10       |
|   &orderby=date      |
|   &order=desc        |
+----------+----------+
           |
           |  Returns: list of recent articles (JSON)
           v
+----------+----------+
|   CHANGE DETECTOR    |
|   Compare article    |
|   IDs against last   |
|   known ID stored    |
|   in SharePoint      |
+----------+----------+
           |
           |  New articles only (0 to N items)
           v
+----------+----------+      For EACH new article:
|   ARTICLE SCRAPER    | <----------------------------------+
|   For each new       |                                    |
|   article:           |     +---------------------------+  |
|   1. GET /knowledge/ |     | Fetch embedded data:      |  |
|      {id}?_embed     |     |  - Author name + title    |  |
|   2. Parse author,   |     |  - Full body text         |  |
|      body, PDFs,     |     |  - PDF links              |  |
|      external links  |     |  - External pub URLs      |  |
|                      |     |  - Featured image         |  |
+----------+----------+     +---------------------------+  |
           |                                                |
           |  Structured article data                       |
           v                                                |
+----------+----------+                                     |
|   AI PROCESSOR       |                                    |
|   Send article text  |                                    |
|   to Azure OpenAI    |                                    |
|   (GPT-4o or GPT-4)  |                                    |
|                      |                                    |
|   Generate:          |                                    |
|   - Summary          |                                    |
|   - Q&A              |                                    |
|   - Keywords/Tags    |                                    |
|   - Audience         |                                    |
|   - Industry         |                                    |
|   - Geography        |                                    |
|   - Solutions        |                                    |
|   - Technology/AI    |                                    |
|   - BD Email Draft   |                                    |
+----------+----------+                                    |
           |                                                |
           |  All 17 column values                          |
           v                                                |
+----------+----------+                                     |
|   SHAREPOINT WRITER  |                                    |
|   Add new row to     |                                    |
|   Excel spreadsheet  |                                    |
|   via Microsoft      |                                    |
|   Graph API          |                                    |
+----------+----------+                                     |
           |                                                |
           |  Row written successfully                      |
           +---> Loop back for next article ----------------+
           |
           v
+----------+----------+
|   UPDATE TRACKER     |
|   Store the latest   |
|   article ID in      |
|   SharePoint list    |
|   (so we don't       |
|   process it again)  |
+----------------------+
```

### Component Interaction Map

```
                     +-------------------+
                     |  accordion.com    |
                     |  WordPress Site   |
                     +--------+----------+
                              |
                    WP REST API (HTTPS GET)
                              |
                              v
+----------------+   +-------+--------+   +------------------+
| Power Automate |-->| Azure Function  |-->| Azure OpenAI     |
| (Scheduler)    |   | (Python 3.11+) |   | (GPT-4o)         |
+----------------+   +-------+--------+   +------------------+
                              |
                    Microsoft Graph API
                              |
                              v
                     +--------+----------+
                     | SharePoint Online |
                     | Excel Workbook    |
                     +-------------------+
```

### What Happens on Each Run (Plain English)

1. Every 6 hours, the automation wakes up.
2. It asks the Accordion website: "What are your 10 most recent knowledge articles?"
3. The website responds with a list of articles in JSON format.
4. The automation checks: "Have I already processed the first article in this list?" It does this by comparing the article's ID to the last ID it remembers processing.
5. If the answer is "yes, I have already seen this one," it stops. No new content.
6. If the answer is "no, this is new," it figures out exactly how many new articles there are (could be 1, could be 5).
7. For each new article, it fetches the full article data (including embedded author info, body text, etc.).
8. It then scrapes additional details from the HTML content: PDF download links, external publication URLs, author names and titles.
9. It sends the article text to Azure OpenAI with a carefully crafted prompt, asking the AI to generate all the enrichment fields (summary, Q&A, keywords, etc.).
10. It takes all the data (both scraped and AI-generated) and writes a new row to the SharePoint Excel spreadsheet.
11. After processing all new articles, it updates its memory of the last processed article ID.

---

## 3. Component Breakdown

### 3.1 Trigger

**Purpose:** Wake up the automation on a regular schedule.

**Two trigger options:**

| Trigger Type | How It Works | Best For |
|---|---|---|
| Power Automate Recurrence | A Power Automate flow with a "Recurrence" trigger set to every 6 hours. The flow calls the Azure Function via HTTP. | Teams already using Power Automate; non-developers. |
| Azure Functions Timer Trigger | A CRON expression inside the Azure Function itself: `0 0 */6 * * *` (every 6 hours at minute 0). | Developers who want everything in one place. |

**Why every 6 hours?**

Accordion publishes thought leadership content during business hours, typically a few times per week. Checking every 6 hours means:
- Content published at 9 AM is captured by 3 PM the same day (worst case).
- We are not hammering the API (only 4 calls per day).
- The WordPress REST API has no published rate limit, but being a good citizen matters.

You can adjust the frequency. Every hour is fine too, but every 6 hours is a good balance.

### 3.2 Data Fetcher

**Purpose:** Call the WordPress REST API to get the most recent articles.

**Endpoint:**
```
GET https://www.accordion.com/wp-json/wp/v2/knowledge?per_page=10&orderby=date&order=desc
```

**What this URL means, piece by piece:**
- `https://www.accordion.com` -- The Accordion website.
- `/wp-json/wp/v2/` -- The standard WordPress REST API prefix.
- `knowledge` -- The custom post type that Accordion uses for their insights/articles (instead of the default `posts`).
- `?per_page=10` -- Return 10 results at a time (the 10 most recent).
- `&orderby=date` -- Sort by publication date.
- `&order=desc` -- Newest first.

**What comes back:** A JSON array of article objects. Each object contains:

| Field | Type | Description |
|---|---|---|
| `id` | Integer | Unique article ID (e.g., `48291`). This is our primary tracking key. |
| `date` | String | ISO 8601 datetime (e.g., `"2026-02-09T14:30:00"`). |
| `slug` | String | URL-friendly title (e.g., `"ai-in-private-equity-due-diligence"`). |
| `title.rendered` | String | Human-readable title with HTML entities (e.g., `"AI in Private Equity Due Diligence"`). |
| `link` | String | Full URL to the article (e.g., `"https://www.accordion.com/knowledge/ai-in-private-equity-due-diligence/"`). |
| `knowledge_type` | Array of Integers | Taxonomy IDs for content type (e.g., `[44]` = Articles). |
| `topics` | Array of Integers | Taxonomy IDs for topics (e.g., `[69, 54]` = AI + Data & Analytics). |
| `company-news` | Array of Integers | Taxonomy IDs for company news categories. |
| `featured_media` | Integer | ID of the featured image attachment. |
| `content.rendered` | String | Full HTML body of the article. |

### 3.3 Change Detector

**Purpose:** Determine which articles are new (haven't been processed yet).

**How it works:**

1. Before calling the API, the automation reads a small SharePoint list (or a single-cell Excel file) that stores one value: `last_processed_id`.
2. After getting the API response, it compares each article's `id` to `last_processed_id`.
3. Any article with an `id` greater than `last_processed_id` is considered new.
4. The articles are processed in chronological order (oldest new article first), so the spreadsheet rows end up in date order.

**Why use article ID instead of date?**

Article IDs in WordPress are auto-incrementing integers. A higher ID always means a newer article. This is simpler and more reliable than comparing dates (which can have timezone issues or be manually backdated).

**Edge case:** If WordPress IDs are not strictly sequential for the `knowledge` post type (because other post types share the same ID sequence), we handle this by storing ALL processed IDs, not just the highest one. See the implementation sections for details.

**Storage location:** A SharePoint list named `AutomationState` with columns:
- `Key` (text): `"last_processed_id"`
- `Value` (text): `"48291"`
- `UpdatedAt` (datetime): Auto-populated

### 3.4 Article Scraper

**Purpose:** For each new article, fetch the complete data including author info, body text, PDF links, and external publication URLs.

**Primary data source -- the `_embed` endpoint:**
```
GET https://www.accordion.com/wp-json/wp/v2/knowledge/{id}?_embed
```

The `?_embed` parameter tells WordPress to include related data inline (author, featured image, taxonomy term names) instead of just IDs. This saves us from making multiple additional API calls.

**What `_embed` gives us:**
- `_embedded.wp:term` -- Full taxonomy term objects (so we get the name "Articles" instead of just the ID `44`).
- `_embedded.wp:featuredmedia` -- The featured image URL and alt text.
- `_embedded.author` -- Author user objects (if the author is a WordPress user).

**What we still need to scrape from the HTML:**

The API's `content.rendered` field gives us the full HTML body. We parse this HTML to extract:

1. **Author name and title:** Found in the "Meet the Author" section. The HTML typically looks like:
   ```html
   <div class="author-bio">
     <h3>Meet the Author</h3>
     <a href="https://www.accordion.com/team/john-smith/">
       <strong>John Smith</strong>
     </a>
     <p>Managing Director, Digital Finance</p>
   </div>
   ```

2. **PDF links:** Found as `<a>` tags with `href` ending in `.pdf`:
   ```html
   <a href="https://www.accordion.com/wp-content/uploads/2026/02/whitepaper.pdf">
     Download the full white paper
   </a>
   ```

3. **External publication URLs:** Articles that originally ran in Forbes, Fortune, CFO Dive, etc. often have a link or mention near the top:
   ```html
   <p>This article originally appeared in
     <a href="https://www.forbes.com/...">Forbes</a>.
   </p>
   ```

4. **FAQ sections:** Some articles have structured FAQ content:
   ```html
   <div class="faq-section">
     <h3>Q: What is value creation in PE?</h3>
     <p>A: Value creation refers to...</p>
   </div>
   ```

**Python libraries used for HTML parsing:**
- `beautifulsoup4` (BeautifulSoup) -- The standard Python HTML parser.
- `re` (regex) -- For pattern matching within text.

### 3.5 AI Processor

**Purpose:** Send the article's full text to an AI model to generate enrichment fields that would take a human 15-30 minutes to write manually.

**Model:** Azure OpenAI GPT-4o (recommended) or GPT-4.

**Why Azure OpenAI instead of direct OpenAI?**
- Data stays within Microsoft's Azure infrastructure (important for enterprise compliance).
- Integrates with Azure Active Directory for authentication.
- Same model quality, different hosting.

**What we send to the AI:**
- The article's title.
- The article's full body text (stripped of HTML tags).
- The article's type (e.g., "Article", "White Paper").
- The article's topics (e.g., "Artificial Intelligence, Data & Analytics").
- The article's author name and title.

**What we ask the AI to generate (9 fields):**

| # | Field | Description | Example Output |
|---|---|---|---|
| 1 | Summary | 2-3 sentence executive summary | "This article explores how PE firms are leveraging AI-driven due diligence tools to accelerate deal evaluation timelines..." |
| 2 | Q&A | 3-5 question-and-answer pairs based on the content | "Q: How does AI improve due diligence? A: AI automates document review..." |
| 3 | Audience | Who should read this (CFOs, PE Partners, Portfolio Company Executives, etc.) | "PE Partners, Operating Partners, Portfolio Company CFOs" |
| 4 | Industry | Industry vertical(s) relevant to the content | "Private Equity, Financial Services" |
| 5 | Geography | Geographic relevance (if mentioned) | "North America" or "Global" |
| 6 | Solutions/Value Creation Levers | Which of Accordion's service areas apply | "Digital Finance, Data & Analytics, Performance Acceleration" |
| 7 | Technology/AI | Technology themes discussed | "Machine Learning, Natural Language Processing, Predictive Analytics" |
| 8 | Keywords/Tags | 5-10 keyword tags including practice names | "AI, due diligence, private equity, deal evaluation, portfolio optimization, Digital Finance" |
| 9 | BD Email Language | A 3-4 sentence client-facing email paragraph | "I wanted to share a recent article from our team that explores how AI is transforming due diligence..." |

The full prompt template is in [Section 9: AI Prompt Engineering](#9-ai-prompt-engineering).

### 3.6 SharePoint Writer

**Purpose:** Write the completed row to the Excel spreadsheet in SharePoint.

**How it works:**

The Microsoft Graph API allows programmatic access to Excel files stored in SharePoint. The specific API call is:

```
POST https://graph.microsoft.com/v1.0/sites/{site-id}/drive/items/{file-id}/workbook/tables/{table-name}/rows/add
```

**Prerequisites:**
1. The Excel file must exist in SharePoint.
2. The data must be formatted as an Excel Table (not just a range of cells). In Excel, select your header row, go to Insert > Table.
3. The table columns must match the 17 columns listed in Section 5.
4. An Azure App Registration must exist with `Sites.ReadWrite.All` permission for the Microsoft Graph API.

**Authentication:**

The Azure Function authenticates to Microsoft Graph using an **App Registration** with a **Client Secret** (or Certificate). This is a service-to-service authentication -- no human login required.

Steps (one-time setup):
1. Go to Azure Portal > Azure Active Directory > App Registrations > New Registration.
2. Name it `AccordionAutomation`.
3. Under API Permissions, add `Microsoft Graph > Application Permissions > Sites.ReadWrite.All`.
4. Grant admin consent.
5. Under Certificates & Secrets, create a new Client Secret. Copy the value.
6. Store the Client ID, Tenant ID, and Client Secret as environment variables in the Azure Function.

---

## 4. Data Flow with Example Payloads

This section shows the exact data that moves between each component, with realistic example values.

### Step 1: Trigger fires, calls Data Fetcher

**Input:** Nothing (the trigger just fires on schedule).

**Action:** HTTP GET request to WordPress.

```
GET https://www.accordion.com/wp-json/wp/v2/knowledge?per_page=10&orderby=date&order=desc
```

**Output:** JSON array (truncated to show 2 articles):

```json
[
  {
    "id": 48295,
    "date": "2026-02-12T09:00:00",
    "date_gmt": "2026-02-12T14:00:00",
    "slug": "ai-transforming-pe-due-diligence",
    "status": "publish",
    "type": "knowledge",
    "link": "https://www.accordion.com/knowledge/ai-transforming-pe-due-diligence/",
    "title": {
      "rendered": "How AI Is Transforming PE Due Diligence"
    },
    "content": {
      "rendered": "<p>In the fast-paced world of private equity, due diligence has traditionally been a labor-intensive process...</p><!-- more HTML content -->",
      "protected": false
    },
    "featured_media": 48296,
    "knowledge_type": [44],
    "topics": [69, 54],
    "company-news": []
  },
  {
    "id": 48290,
    "date": "2026-02-09T10:00:00",
    "date_gmt": "2026-02-09T15:00:00",
    "slug": "cfo-guide-to-ai-implementation",
    "status": "publish",
    "type": "knowledge",
    "link": "https://www.accordion.com/knowledge/cfo-guide-to-ai-implementation/",
    "title": {
      "rendered": "The CFO&#8217;s Guide to AI Implementation"
    },
    "content": {
      "rendered": "<p>As artificial intelligence tools become more accessible...</p>",
      "protected": false
    },
    "featured_media": 48291,
    "knowledge_type": [44],
    "topics": [69, 59],
    "company-news": []
  }
]
```

### Step 2: Change Detector compares IDs

**Input:**
- API response (array of articles, newest first).
- `last_processed_id` from SharePoint: `48290`.

**Logic:**
```
For each article in API response:
    if article.id > last_processed_id:
        mark as NEW
    else:
        stop (all remaining are older)
```

**Output:** List of new article IDs to process:
```json
{
  "new_article_ids": [48295],
  "last_processed_id_before": 48290,
  "count": 1
}
```

In this example, article `48295` is new (it is greater than `48290`). Article `48290` has already been processed, so we stop.

### Step 3: Article Scraper fetches full data for each new article

**Input:** Article ID `48295`.

**Action:** HTTP GET request:
```
GET https://www.accordion.com/wp-json/wp/v2/knowledge/48295?_embed
```

**Output:** Full article JSON with embedded data (key fields shown):

```json
{
  "id": 48295,
  "date": "2026-02-12T09:00:00",
  "title": {
    "rendered": "How AI Is Transforming PE Due Diligence"
  },
  "content": {
    "rendered": "<p>In the fast-paced world of private equity...</p><h2>The Traditional Approach</h2><p>Historically, due diligence teams would spend weeks...</p><div class='author-bio'><h3>Meet the Author</h3><a href='https://www.accordion.com/team/sarah-chen/'><strong>Sarah Chen</strong></a><p>Managing Director, Digital Finance</p></div><div class='faq-section'><h3>Q: How long does AI-assisted due diligence take?</h3><p>A: With AI tools, the initial screening phase can be reduced from weeks to days...</p></div>"
  },
  "link": "https://www.accordion.com/knowledge/ai-transforming-pe-due-diligence/",
  "knowledge_type": [44],
  "topics": [69, 54],
  "_embedded": {
    "wp:term": [
      [
        {"id": 44, "name": "Articles", "taxonomy": "knowledge_type"}
      ],
      [
        {"id": 69, "name": "Artificial Intelligence", "taxonomy": "topics"},
        {"id": 54, "name": "Data & Analytics", "taxonomy": "topics"}
      ]
    ],
    "wp:featuredmedia": [
      {
        "source_url": "https://www.accordion.com/wp-content/uploads/2026/02/ai-due-diligence.jpg"
      }
    ]
  }
}
```

**After HTML parsing, the structured scrape result:**

```json
{
  "id": 48295,
  "title": "How AI Is Transforming PE Due Diligence",
  "publish_date": "2026-02-12",
  "url": "https://www.accordion.com/knowledge/ai-transforming-pe-due-diligence/",
  "type": "Articles",
  "topics": ["Artificial Intelligence", "Data & Analytics"],
  "authors": [
    {
      "name": "Sarah Chen",
      "title": "Managing Director, Digital Finance",
      "profile_url": "https://www.accordion.com/team/sarah-chen/"
    }
  ],
  "body_text": "In the fast-paced world of private equity, due diligence has traditionally been a labor-intensive process... [full plain text, HTML stripped]",
  "body_html": "<p>In the fast-paced world of private equity...</p>...",
  "pdf_links": [],
  "external_publication": null,
  "external_publication_url": null,
  "faq_content": [
    {
      "question": "How long does AI-assisted due diligence take?",
      "answer": "With AI tools, the initial screening phase can be reduced from weeks to days..."
    }
  ],
  "word_count": 1847
}
```

### Step 4: AI Processor generates enrichment fields

**Input sent to Azure OpenAI:**

```json
{
  "model": "gpt-4o",
  "messages": [
    {
      "role": "system",
      "content": "[System prompt -- see Section 9]"
    },
    {
      "role": "user",
      "content": "TITLE: How AI Is Transforming PE Due Diligence\nTYPE: Articles\nTOPICS: Artificial Intelligence, Data & Analytics\nAUTHOR: Sarah Chen, Managing Director, Digital Finance\nDATE: 2026-02-12\n\nARTICLE TEXT:\nIn the fast-paced world of private equity, due diligence has traditionally been a labor-intensive process... [full body text]"
    }
  ],
  "response_format": {"type": "json_object"},
  "temperature": 0.3,
  "max_tokens": 2000
}
```

**Output from Azure OpenAI:**

```json
{
  "summary": "This article examines how artificial intelligence is reshaping the due diligence process in private equity, reducing timelines from weeks to days. It covers specific AI applications including automated document review, predictive financial modeling, and risk pattern detection, while noting the continued importance of human judgment in final investment decisions.",
  "qa": "Q: How does AI improve the due diligence process?\nA: AI automates document review, identifies risk patterns across large datasets, and generates predictive financial models, reducing the initial screening phase from weeks to days.\n\nQ: What are the limitations of AI in due diligence?\nA: AI excels at data processing but cannot replace human judgment in assessing management quality, cultural fit, and strategic alignment.\n\nQ: What types of AI tools are most commonly used?\nA: Natural language processing for document analysis, machine learning for financial pattern recognition, and predictive analytics for forecasting.",
  "audience": "PE Partners, Operating Partners, Due Diligence Teams, Portfolio Company CFOs",
  "industry": "Private Equity, Financial Services",
  "geography": "Global",
  "solutions_value_creation": "Digital Finance, Data & Analytics, Performance Acceleration",
  "technology_ai": "Artificial Intelligence, Machine Learning, Natural Language Processing, Predictive Analytics, Document Automation",
  "keywords_tags": "AI, due diligence, private equity, deal evaluation, document review, predictive modeling, risk assessment, Digital Finance, Data & Analytics",
  "bd_email_language": "I wanted to share a recent article from our Digital Finance team that explores how AI is fundamentally changing the due diligence process in private equity. The piece examines how leading firms are using AI-driven tools to reduce screening timelines from weeks to days while improving the quality of their analysis. Given the increasing pace of deal flow, I thought this might be relevant to your team's evaluation processes."
}
```

### Step 5: SharePoint Writer assembles and writes the row

**Input:** Combined scraped data + AI-generated data.

**The final row (all 17 columns):**

```json
{
  "values": [
    [
      "How AI Is Transforming PE Due Diligence",
      "Articles",
      "This article examines how artificial intelligence is reshaping the due diligence process in private equity, reducing timelines from weeks to days. It covers specific AI applications including automated document review, predictive financial modeling, and risk pattern detection, while noting the continued importance of human judgment in final investment decisions.",
      "Q: How does AI improve the due diligence process?\nA: AI automates document review, identifies risk patterns across large datasets, and generates predictive financial models, reducing the initial screening phase from weeks to days.\n\nQ: What are the limitations of AI in due diligence?\nA: AI excels at data processing but cannot replace human judgment in assessing management quality, cultural fit, and strategic alignment.\n\nQ: What types of AI tools are most commonly used?\nA: Natural language processing for document analysis, machine learning for financial pattern recognition, and predictive analytics for forecasting.",
      "Sarah Chen, Managing Director, Digital Finance",
      "2026-02-12",
      "",
      "https://www.accordion.com/knowledge/ai-transforming-pe-due-diligence/",
      "",
      "",
      "PE Partners, Operating Partners, Due Diligence Teams, Portfolio Company CFOs",
      "Private Equity, Financial Services",
      "Global",
      "Digital Finance, Data & Analytics, Performance Acceleration",
      "Artificial Intelligence, Machine Learning, Natural Language Processing, Predictive Analytics, Document Automation",
      "AI, due diligence, private equity, deal evaluation, document review, predictive modeling, risk assessment, Digital Finance, Data & Analytics",
      "I wanted to share a recent article from our Digital Finance team that explores how AI is fundamentally changing the due diligence process in private equity. The piece examines how leading firms are using AI-driven tools to reduce screening timelines from weeks to days while improving the quality of their analysis. Given the increasing pace of deal flow, I thought this might be relevant to your team's evaluation processes."
    ]
  ]
}
```

**API call to Microsoft Graph:**

```
POST https://graph.microsoft.com/v1.0/sites/{site-id}/drive/items/{file-id}/workbook/tables/ThoughtLeadership/rows/add
Authorization: Bearer {access-token}
Content-Type: application/json

{
  "values": [[ ...the 17 values above... ]]
}
```

### Step 6: Update the tracker

**Action:** Update the SharePoint list item:

```
PATCH https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-id}/items/{item-id}/fields
Authorization: Bearer {access-token}
Content-Type: application/json

{
  "Value": "48295",
  "UpdatedAt": "2026-02-12T15:00:00Z"
}
```

---

## 5. Column Mapping Logic (All 17 Columns)

This section explains, for every single column in the SharePoint spreadsheet, exactly where the data comes from.

### Legend

| Source Type | Meaning |
|---|---|
| **API-Direct** | Comes directly from the WordPress REST API response, no transformation needed. |
| **API-Mapped** | Comes from the API but needs a lookup/transformation (e.g., taxonomy ID to name). |
| **HTML-Scraped** | Parsed from the article's HTML content (`content.rendered`). |
| **AI-Generated** | Produced by the Azure OpenAI model based on the article text. |
| **Conditional** | Only populated when certain conditions are met (e.g., external publications). |

---

### Column 1: Topic Title

| Attribute | Value |
|---|---|
| **Source** | API-Direct |
| **API Field** | `title.rendered` |
| **Transformation** | Decode HTML entities (`&#8217;` becomes `'`, `&amp;` becomes `&`). |
| **Example Input** | `"The CFO&#8217;s Guide to AI Implementation"` |
| **Example Output** | `"The CFO's Guide to AI Implementation"` |
| **Python Code** | `html.unescape(article["title"]["rendered"])` |
| **Can it be empty?** | No. Every article has a title. |

---

### Column 2: Type

| Attribute | Value |
|---|---|
| **Source** | API-Mapped |
| **API Field** | `knowledge_type` (array of taxonomy IDs) |
| **Transformation** | Map the numeric ID to its human-readable name using the taxonomy lookup table. |
| **Taxonomy Lookup Table** | |

```python
KNOWLEDGE_TYPE_MAP = {
    44: "Articles",
    46: "Event Recaps",
    49: "Multimedia",
    48: "Press Releases",
    47: "White Papers"
}
```

| **Example Input** | `[44]` |
|---|---|
| **Example Output** | `"Articles"` |
| **Python Code** | `KNOWLEDGE_TYPE_MAP.get(article["knowledge_type"][0], "Unknown")` |
| **Can it be empty?** | Rarely. If a new type is added to WordPress, it would show as "Unknown." |

---

### Column 3: Summary

| Attribute | Value |
|---|---|
| **Source** | AI-Generated |
| **Input to AI** | Full article body text (HTML stripped). |
| **Prompt Instruction** | "Write a 2-3 sentence executive summary of this article. Focus on the key takeaway and who would benefit from reading it." |
| **Example Output** | "This article examines how artificial intelligence is reshaping the due diligence process in private equity, reducing timelines from weeks to days..." |
| **Can it be empty?** | Only if AI generation fails (see Error Handling). |
| **Fallback** | First 200 characters of the article body text + "..." |

---

### Column 4: Q&A

| Attribute | Value |
|---|---|
| **Source** | AI-Generated (supplemented by HTML-Scraped FAQ if present) |
| **Input to AI** | Full article body text. If the article already contains FAQ sections, those are included as context. |
| **Prompt Instruction** | "Generate 3-5 Q&A pairs based on this article. If the article contains its own FAQ section, incorporate those questions. Format as Q: [question]\nA: [answer]" |
| **Example Output** | "Q: How does AI improve due diligence?\nA: AI automates document review..." |
| **Can it be empty?** | Only if AI generation fails. |
| **Fallback** | If the article has a native FAQ section, use that. Otherwise, leave blank with a note "AI generation failed." |

---

### Column 5: Authors

| Attribute | Value |
|---|---|
| **Source** | HTML-Scraped |
| **HTML Location** | The "Meet the Author" section in `content.rendered`. |
| **Parsing Logic** | 1. Find `<div>` with class containing "author". 2. Extract the `<strong>` tag text (author name). 3. Extract the following `<p>` tag text (author title). 4. If multiple authors, join with semicolons. |
| **Example Input HTML** | `<a href="/team/sarah-chen/"><strong>Sarah Chen</strong></a><p>Managing Director, Digital Finance</p>` |
| **Example Output** | `"Sarah Chen, Managing Director, Digital Finance"` |
| **Multiple Authors** | `"Sarah Chen, Managing Director, Digital Finance; James Park, Senior Director, Data & Analytics"` |
| **Can it be empty?** | Yes. Some content (especially press releases or multimedia) may not have a listed author. |
| **Fallback** | Check the `_embedded.author` field from the API (WordPress user who published the post). If neither exists, leave blank. |

---

### Column 6: Publish Date

| Attribute | Value |
|---|---|
| **Source** | API-Direct |
| **API Field** | `date` |
| **Transformation** | Parse ISO 8601 datetime and format as `YYYY-MM-DD` (or whatever date format the SharePoint column expects). |
| **Example Input** | `"2026-02-12T09:00:00"` |
| **Example Output** | `"2026-02-12"` |
| **Python Code** | `datetime.fromisoformat(article["date"]).strftime("%Y-%m-%d")` |
| **Can it be empty?** | No. Every article has a publish date. |

---

### Column 7: Publication

| Attribute | Value |
|---|---|
| **Source** | HTML-Scraped + Conditional |
| **When it applies** | Only for articles that originally appeared in an external publication (Forbes, Fortune, CFO Dive, etc.). |
| **Parsing Logic** | 1. Search the body HTML for patterns like "originally appeared in", "as published in", "featured in". 2. Extract the publication name from the surrounding `<a>` tag or text. |
| **Example Input HTML** | `<p>This article originally appeared in <a href="https://www.forbes.com/...">Forbes</a>.</p>` |
| **Example Output** | `"Forbes"` |
| **Can it be empty?** | Yes. Most articles are Accordion-original and have no external publication. |
| **Fallback** | Leave blank. |

---

### Column 8: URL

| Attribute | Value |
|---|---|
| **Source** | API-Direct |
| **API Field** | `link` |
| **Transformation** | None. Use as-is. |
| **Example Output** | `"https://www.accordion.com/knowledge/ai-transforming-pe-due-diligence/"` |
| **Can it be empty?** | No. Every article has a URL. |

---

### Column 9: Link to PDF

| Attribute | Value |
|---|---|
| **Source** | HTML-Scraped + Conditional |
| **When it applies** | Only for articles/white papers that include a downloadable PDF. |
| **Parsing Logic** | 1. Search all `<a>` tags in `content.rendered`. 2. Filter for `href` attributes ending in `.pdf`. 3. If multiple PDFs, join with semicolons. |
| **Example Input HTML** | `<a href="https://www.accordion.com/wp-content/uploads/2026/02/whitepaper.pdf">Download PDF</a>` |
| **Example Output** | `"https://www.accordion.com/wp-content/uploads/2026/02/whitepaper.pdf"` |
| **Can it be empty?** | Yes. Most articles do not have PDFs. White papers are more likely to. |
| **Fallback** | Leave blank. |

---

### Column 10: Publication URL

| Attribute | Value |
|---|---|
| **Source** | HTML-Scraped + Conditional |
| **When it applies** | Only when Column 7 (Publication) is populated. This is the actual URL to the article on the external publication's site. |
| **Parsing Logic** | Same detection as Column 7 (Publication). Extract the `href` from the `<a>` tag that links to the external site. |
| **Example Output** | `"https://www.forbes.com/sites/forbesfinancecouncil/2026/02/10/ai-due-diligence/"` |
| **Can it be empty?** | Yes. Same conditions as Column 7. |
| **Fallback** | Leave blank. |

---

### Column 11: Audience

| Attribute | Value |
|---|---|
| **Source** | AI-Generated |
| **Input to AI** | Article body text + type + topics. |
| **Prompt Instruction** | "Identify the target audience for this article. Use labels from this list where applicable: PE Partners, Operating Partners, Portfolio Company CFOs, Portfolio Company CEOs, Finance Teams, IT Leaders, Board Members, Sponsors, LPs. You may add other specific audiences if the content warrants it. Provide as a comma-separated list." |
| **Example Output** | `"PE Partners, Operating Partners, Due Diligence Teams, Portfolio Company CFOs"` |
| **Can it be empty?** | Only if AI generation fails. |
| **Fallback** | Default to the content type's typical audience: Articles -> "PE Partners, Portfolio Company CFOs"; White Papers -> "PE Partners, Operating Partners, Finance Teams". |

---

### Column 12: Industry

| Attribute | Value |
|---|---|
| **Source** | AI-Generated |
| **Input to AI** | Article body text. |
| **Prompt Instruction** | "Identify the industry vertical(s) this article is relevant to. Use labels from this list where applicable: Private Equity, Financial Services, Healthcare, Technology, Manufacturing, Retail, Energy. If the article is broadly applicable, use 'Cross-Industry'. Provide as a comma-separated list." |
| **Example Output** | `"Private Equity, Financial Services"` |
| **Can it be empty?** | Only if AI generation fails. |
| **Fallback** | `"Private Equity"` (since all Accordion content is PE-focused). |

---

### Column 13: Geography

| Attribute | Value |
|---|---|
| **Source** | AI-Generated |
| **Input to AI** | Article body text. |
| **Prompt Instruction** | "Identify the geographic relevance of this article. If specific regions, countries, or markets are discussed, list them. If the content is broadly applicable, use 'Global'. Options include: Global, North America, Europe, Asia-Pacific, Middle East, Latin America, or specific countries." |
| **Example Output** | `"Global"` or `"North America, Europe"` |
| **Can it be empty?** | Only if AI generation fails. |
| **Fallback** | `"Global"` |

---

### Column 14: Solutions/Value Creation Levers

| Attribute | Value |
|---|---|
| **Source** | AI-Generated (informed by API-Mapped topics) |
| **Input to AI** | Article body text + topics taxonomy. |
| **Prompt Instruction** | "Identify which of Accordion's solutions or value creation levers are discussed in this article. Use labels from this list where applicable: Foundational Accounting, FP&A Enhancement, Digital Finance, Data & Analytics, Performance Acceleration, Exit Planning, Transaction Support, Supply Chain & Operational Logistics. Include the article's listed topics as a starting point. Provide as a comma-separated list." |
| **Example Output** | `"Digital Finance, Data & Analytics, Performance Acceleration"` |
| **Can it be empty?** | Only if AI generation fails. |
| **Fallback** | Map the article's `topics` taxonomy directly using the taxonomy lookup table. |

---

### Column 15: Technology/AI

| Attribute | Value |
|---|---|
| **Source** | AI-Generated |
| **Input to AI** | Article body text. |
| **Prompt Instruction** | "List the specific technologies, AI techniques, or digital tools discussed in this article. Examples: Machine Learning, Natural Language Processing, Robotic Process Automation, Predictive Analytics, Cloud Computing, ERP Systems, Business Intelligence, Generative AI, Large Language Models. If no specific technology is discussed, write 'N/A'. Provide as a comma-separated list." |
| **Example Output** | `"Artificial Intelligence, Machine Learning, Natural Language Processing, Predictive Analytics, Document Automation"` |
| **Can it be empty?** | Yes. Some articles (e.g., general advisory pieces) may not discuss specific technology. Output `"N/A"`. |

---

### Column 16: Keywords/Tags

| Attribute | Value |
|---|---|
| **Source** | AI-Generated (supplemented by API-Mapped topics) |
| **Input to AI** | Article body text + title + topics. |
| **Prompt Instruction** | "Generate 5-10 keyword tags for this article. Include: (a) the article's listed topic names, (b) specific Accordion practice names mentioned, (c) key themes and concepts. Provide as a comma-separated list." |
| **Example Output** | `"AI, due diligence, private equity, deal evaluation, document review, predictive modeling, risk assessment, Digital Finance, Data & Analytics"` |
| **Can it be empty?** | Only if AI generation fails. |
| **Fallback** | Use the topic names from the taxonomy as basic tags. |

---

### Column 17: Business Development Client-Facing Email Language

| Attribute | Value |
|---|---|
| **Source** | AI-Generated |
| **Input to AI** | Article title + summary (generated in Column 3) + author name. |
| **Prompt Instruction** | "Write a 3-4 sentence email paragraph that an Accordion business development professional could use to share this article with a client or prospect. The tone should be professional, consultative, and value-driven. Do not include a greeting or sign-off -- just the body paragraph. Reference the article's key insight and why it would be relevant to the recipient." |
| **Example Output** | `"I wanted to share a recent article from our Digital Finance team that explores how AI is fundamentally changing the due diligence process in private equity..."` |
| **Can it be empty?** | Only if AI generation fails. |
| **Fallback** | A generic template: "I wanted to share a recent [Type] from our team: [Title]. You can read it here: [URL]" |

---

### Summary Table: All 17 Columns

| # | Column Name | Source Type | API Field / Method | AI-Generated? |
|---|---|---|---|---|
| 1 | Topic Title | API-Direct | `title.rendered` | No |
| 2 | Type | API-Mapped | `knowledge_type` -> lookup | No |
| 3 | Summary | AI-Generated | -- | Yes |
| 4 | Q&A | AI-Generated | -- | Yes |
| 5 | Authors | HTML-Scraped | `content.rendered` parse | No |
| 6 | Publish Date | API-Direct | `date` | No |
| 7 | Publication | HTML-Scraped | `content.rendered` parse | No |
| 8 | URL | API-Direct | `link` | No |
| 9 | Link to PDF | HTML-Scraped | `content.rendered` parse | No |
| 10 | Publication URL | HTML-Scraped | `content.rendered` parse | No |
| 11 | Audience | AI-Generated | -- | Yes |
| 12 | Industry | AI-Generated | -- | Yes |
| 13 | Geography | AI-Generated | -- | Yes |
| 14 | Solutions/Value Creation Levers | AI-Generated | -- | Yes |
| 15 | Technology/AI | AI-Generated | -- | Yes |
| 16 | Keywords/Tags | AI-Generated | -- | Yes |
| 17 | BD Email Language | AI-Generated | -- | Yes |

**Totals: 4 API-Direct, 1 API-Mapped, 4 HTML-Scraped, 9 AI-Generated** (some columns use multiple sources).

---

## 6. Implementation Option A: Power Automate + Azure Function

**Best for:** Teams that already use Microsoft 365 and want a low-code approach with visual monitoring.

### Architecture

```
Power Automate Flow
    |
    +--> Recurrence Trigger (every 6 hours)
    |
    +--> HTTP Action: Call Azure Function URL
    |       (passes no parameters; the function handles everything)
    |
    +--> Condition: Check if function returned new articles
    |       |
    |       +--> Yes: Log success, optionally send Teams notification
    |       |
    |       +--> No: Log "no new articles"
    |
    +--> (Azure Function handles all logic internally)
```

### Step-by-Step Setup

#### Part 1: Create the Azure Function (Python)

1. **Go to the Azure Portal** (portal.azure.com).
2. **Create a Function App:**
   - Search for "Function App" and click Create.
   - Subscription: Your Azure subscription.
   - Resource Group: Create a new one called `rg-accordion-automation`.
   - Function App Name: `func-accordion-insights` (must be globally unique).
   - Runtime Stack: Python 3.11.
   - Region: East US (or whatever is closest to your team).
   - Plan: Consumption (Serverless) -- you only pay for execution time.
   - Click Review + Create, then Create.

3. **Create the function within the app:**
   - Once deployed, go to your Function App.
   - Click Functions > Create.
   - Template: HTTP Trigger.
   - Name: `ProcessNewArticles`.
   - Authorization Level: Function (requires a key to call it).

4. **Add the Python code:**
   - The complete Python code is below. This goes in `__init__.py`.

5. **Add application settings (environment variables):**
   - Go to Configuration > Application Settings.
   - Add the following:

| Setting Name | Value | Description |
|---|---|---|
| `AZURE_OPENAI_ENDPOINT` | `https://your-instance.openai.azure.com/` | Your Azure OpenAI endpoint |
| `AZURE_OPENAI_KEY` | `abc123...` | Your Azure OpenAI API key |
| `AZURE_OPENAI_DEPLOYMENT` | `gpt-4o` | Your deployed model name |
| `GRAPH_CLIENT_ID` | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` | Azure App Registration Client ID |
| `GRAPH_CLIENT_SECRET` | `your-secret-value` | Azure App Registration Client Secret |
| `GRAPH_TENANT_ID` | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` | Azure AD Tenant ID |
| `SHAREPOINT_SITE_ID` | `your-site-id` | SharePoint site ID |
| `SHAREPOINT_DRIVE_ID` | `your-drive-id` | SharePoint document library drive ID |
| `SHAREPOINT_FILE_ID` | `your-file-id` | The Excel file's item ID |
| `SHAREPOINT_TABLE_NAME` | `ThoughtLeadership` | The Excel table name |
| `SHAREPOINT_LIST_ID` | `your-list-id` | The tracking list ID |

6. **Add `requirements.txt`:**

```
azure-functions
requests
beautifulsoup4
openai
msal
```

#### Complete Python Code for Azure Function

```python
"""
Azure Function: ProcessNewArticles
Monitors Accordion's WordPress API for new knowledge articles,
enriches them with AI-generated metadata, and writes to SharePoint Excel.
"""

import azure.functions as func
import json
import logging
import os
import re
import html as html_module
from datetime import datetime
from typing import Optional

import requests
from bs4 import BeautifulSoup
from openai import AzureOpenAI
import msal

# ---------------------------------------------------------------------------
# Configuration (loaded from environment variables)
# ---------------------------------------------------------------------------
WP_API_BASE = "https://www.accordion.com/wp-json/wp/v2"
WP_KNOWLEDGE_ENDPOINT = f"{WP_API_BASE}/knowledge"

AZURE_OPENAI_ENDPOINT = os.environ["AZURE_OPENAI_ENDPOINT"]
AZURE_OPENAI_KEY = os.environ["AZURE_OPENAI_KEY"]
AZURE_OPENAI_DEPLOYMENT = os.environ["AZURE_OPENAI_DEPLOYMENT"]

GRAPH_CLIENT_ID = os.environ["GRAPH_CLIENT_ID"]
GRAPH_CLIENT_SECRET = os.environ["GRAPH_CLIENT_SECRET"]
GRAPH_TENANT_ID = os.environ["GRAPH_TENANT_ID"]

SHAREPOINT_SITE_ID = os.environ["SHAREPOINT_SITE_ID"]
SHAREPOINT_DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]
SHAREPOINT_FILE_ID = os.environ["SHAREPOINT_FILE_ID"]
SHAREPOINT_TABLE_NAME = os.environ["SHAREPOINT_TABLE_NAME"]
SHAREPOINT_LIST_ID = os.environ["SHAREPOINT_LIST_ID"]

# ---------------------------------------------------------------------------
# Taxonomy mappings
# ---------------------------------------------------------------------------
KNOWLEDGE_TYPE_MAP = {
    44: "Articles",
    46: "Event Recaps",
    49: "Multimedia",
    48: "Press Releases",
    47: "White Papers",
}

TOPICS_MAP = {
    52: "Advice for Sponsors & CFOs",
    69: "Artificial Intelligence",
    54: "Data & Analytics",
    59: "Digital Finance",
    63: "Exit Planning and Transaction Support",
    57: "Foundational Accounting and FP&A Enhancement",
    67: "Healthcare",
    58: "Performance Acceleration",
    68: "Supply Chain & Operational Logistics",
    66: "Tech Tutorials",
}


# ---------------------------------------------------------------------------
# Microsoft Graph authentication
# ---------------------------------------------------------------------------
def get_graph_token() -> str:
    """
    Authenticate with Microsoft Graph using client credentials (app-only).
    Returns a Bearer access token.
    """
    authority = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        GRAPH_CLIENT_ID,
        authority=authority,
        client_credential=GRAPH_CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed to acquire Graph token: {result.get('error_description')}")


# ---------------------------------------------------------------------------
# SharePoint helpers: read/write tracking state
# ---------------------------------------------------------------------------
def get_last_processed_id(token: str) -> int:
    """
    Read the last processed article ID from the SharePoint tracking list.
    Returns 0 if no value is found (first run).
    """
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
        f"/lists/{SHAREPOINT_LIST_ID}/items"
        f"?$filter=fields/Key eq 'last_processed_id'"
        f"&$expand=fields"
    )
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    items = resp.json().get("value", [])
    if items:
        return int(items[0]["fields"].get("Value", "0"))
    return 0


def update_last_processed_id(token: str, article_id: int) -> None:
    """
    Update (or create) the last processed article ID in the SharePoint
    tracking list.
    """
    # First, try to find the existing item
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
        f"/lists/{SHAREPOINT_LIST_ID}/items"
        f"?$filter=fields/Key eq 'last_processed_id'"
        f"&$expand=fields"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    items = resp.json().get("value", [])

    if items:
        # Update existing item
        item_id = items[0]["id"]
        patch_url = (
            f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
            f"/lists/{SHAREPOINT_LIST_ID}/items/{item_id}/fields"
        )
        requests.patch(
            patch_url,
            headers=headers,
            json={"Value": str(article_id)},
        ).raise_for_status()
    else:
        # Create new item
        post_url = (
            f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
            f"/lists/{SHAREPOINT_LIST_ID}/items"
        )
        requests.post(
            post_url,
            headers=headers,
            json={"fields": {"Key": "last_processed_id", "Value": str(article_id)}},
        ).raise_for_status()


# ---------------------------------------------------------------------------
# WordPress API: fetch articles
# ---------------------------------------------------------------------------
def fetch_recent_articles(count: int = 10) -> list[dict]:
    """
    Fetch the N most recent knowledge articles from the WordPress REST API.
    Returns a list of article dicts, newest first.
    """
    params = {
        "per_page": count,
        "orderby": "date",
        "order": "desc",
    }
    resp = requests.get(WP_KNOWLEDGE_ENDPOINT, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def fetch_article_detail(article_id: int) -> dict:
    """
    Fetch a single article with embedded data (author, taxonomy terms,
    featured media).
    """
    url = f"{WP_KNOWLEDGE_ENDPOINT}/{article_id}?_embed"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# HTML parsing: extract structured data from article content
# ---------------------------------------------------------------------------
def parse_article_html(article: dict) -> dict:
    """
    Parse the article's HTML content to extract:
    - Author name(s) and title(s)
    - PDF links
    - External publication name and URL
    - FAQ content
    - Plain body text
    """
    content_html = article.get("content", {}).get("rendered", "")
    soup = BeautifulSoup(content_html, "html.parser")

    # --- Extract authors from "Meet the Author" section ---
    authors = []
    # Strategy 1: Look for author-bio div or similar class
    author_sections = soup.find_all(
        ["div", "section"],
        class_=re.compile(r"author", re.IGNORECASE),
    )
    for section in author_sections:
        name_tag = section.find("strong")
        if name_tag:
            name = name_tag.get_text(strip=True)
            # The title is often in the next <p> tag after the name
            title_tag = section.find("p")
            title = title_tag.get_text(strip=True) if title_tag else ""
            profile_link = ""
            link_tag = section.find("a", href=re.compile(r"/team/"))
            if link_tag:
                profile_link = link_tag["href"]
            authors.append({
                "name": name,
                "title": title,
                "profile_url": profile_link,
            })

    # Strategy 2: If no author-bio div, look for "Meet the Author" heading
    if not authors:
        for heading in soup.find_all(["h2", "h3", "h4"]):
            if "meet the author" in heading.get_text(strip=True).lower():
                # The author info is in the siblings after this heading
                for sibling in heading.find_next_siblings():
                    name_tag = sibling.find("strong")
                    if name_tag:
                        name = name_tag.get_text(strip=True)
                        title_tag = sibling.find("p")
                        title = title_tag.get_text(strip=True) if title_tag else ""
                        authors.append({"name": name, "title": title, "profile_url": ""})

    # Strategy 3: Check _embedded author data from WP API
    if not authors and "_embedded" in article:
        embedded_authors = article["_embedded"].get("author", [])
        for ea in embedded_authors:
            authors.append({
                "name": ea.get("name", ""),
                "title": "",
                "profile_url": "",
            })

    # --- Extract PDF links ---
    pdf_links = []
    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"]
        if href.lower().endswith(".pdf"):
            pdf_links.append(href)

    # --- Extract external publication info ---
    external_pub_name = None
    external_pub_url = None
    pub_patterns = [
        r"originally\s+(?:appeared|published)\s+(?:in|on)\s+",
        r"as\s+(?:seen|published|featured)\s+(?:in|on)\s+",
        r"featured\s+(?:in|on)\s+",
    ]
    body_text_raw = soup.get_text(" ", strip=True)
    for pattern in pub_patterns:
        match = re.search(pattern, body_text_raw, re.IGNORECASE)
        if match:
            # Find the <a> tag closest to this text
            for a_tag in soup.find_all("a", href=True):
                parent_text = a_tag.parent.get_text(" ", strip=True) if a_tag.parent else ""
                if re.search(pattern, parent_text, re.IGNORECASE):
                    external_pub_url = a_tag["href"]
                    external_pub_name = a_tag.get_text(strip=True)
                    break
            if external_pub_name:
                break

    # --- Extract FAQ content ---
    faq_items = []
    # Look for FAQ sections by class
    faq_sections = soup.find_all(
        ["div", "section"],
        class_=re.compile(r"faq", re.IGNORECASE),
    )
    for faq in faq_sections:
        questions = faq.find_all(["h3", "h4", "dt", "strong"])
        for q in questions:
            q_text = q.get_text(strip=True)
            if q_text.startswith("Q:"):
                q_text = q_text[2:].strip()
            # Find the answer (next sibling <p> or <dd>)
            a_tag = q.find_next_sibling(["p", "dd", "div"])
            a_text = a_tag.get_text(strip=True) if a_tag else ""
            if a_text.startswith("A:"):
                a_text = a_text[2:].strip()
            faq_items.append({"question": q_text, "answer": a_text})

    # --- Extract plain body text (strip all HTML) ---
    # Remove script and style elements first
    for tag in soup(["script", "style"]):
        tag.decompose()
    body_text = soup.get_text(" ", strip=True)
    # Clean up excessive whitespace
    body_text = re.sub(r"\s+", " ", body_text).strip()

    return {
        "authors": authors,
        "pdf_links": pdf_links,
        "external_publication": external_pub_name,
        "external_publication_url": external_pub_url,
        "faq_content": faq_items,
        "body_text": body_text,
    }


# ---------------------------------------------------------------------------
# Map taxonomy IDs to human-readable names
# ---------------------------------------------------------------------------
def map_knowledge_type(type_ids: list[int]) -> str:
    """Map knowledge_type taxonomy IDs to names."""
    names = [KNOWLEDGE_TYPE_MAP.get(tid, f"Unknown({tid})") for tid in type_ids]
    return ", ".join(names) if names else "Unknown"


def map_topics(topic_ids: list[int]) -> str:
    """Map topics taxonomy IDs to names."""
    names = [TOPICS_MAP.get(tid, f"Unknown({tid})") for tid in topic_ids]
    return ", ".join(names) if names else ""


# ---------------------------------------------------------------------------
# AI enrichment: call Azure OpenAI
# ---------------------------------------------------------------------------
def generate_ai_enrichment(
    title: str,
    body_text: str,
    content_type: str,
    topics: str,
    author_str: str,
) -> dict:
    """
    Send article content to Azure OpenAI and get back structured enrichment
    data (summary, Q&A, audience, industry, geography, solutions, tech,
    keywords, BD email).

    Returns a dict with keys matching the AI-generated columns.
    """
    client = AzureOpenAI(
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_key=AZURE_OPENAI_KEY,
        api_version="2024-08-01-preview",
    )

    system_prompt = """You are an expert content analyst for Accordion, a private equity-focused
financial consulting firm. Your job is to analyze thought leadership content and generate
structured metadata for a marketing database.

You must return a valid JSON object with exactly these keys:
- summary
- qa
- audience
- industry
- geography
- solutions_value_creation
- technology_ai
- keywords_tags
- bd_email_language

Guidelines for each field:

SUMMARY: Write a 2-3 sentence executive summary. Focus on the key takeaway and who benefits.

QA: Generate 3-5 Q&A pairs. Format as "Q: [question]\\nA: [answer]\\n\\n" for each pair.

AUDIENCE: Comma-separated list. Use from: PE Partners, Operating Partners, Portfolio Company CFOs,
Portfolio Company CEOs, Finance Teams, IT Leaders, Board Members, Sponsors, LPs,
Due Diligence Teams, Fund Administrators. Add others if specific to the content.

INDUSTRY: Comma-separated list. Use from: Private Equity, Financial Services, Healthcare,
Technology, Manufacturing, Retail, Energy, Cross-Industry.

GEOGRAPHY: Comma-separated list. Use from: Global, North America, Europe, Asia-Pacific,
Middle East, Latin America, or specific countries mentioned. Default to "Global" if no
specific geography is discussed.

SOLUTIONS_VALUE_CREATION: Comma-separated list. Use from: Foundational Accounting,
FP&A Enhancement, Digital Finance, Data & Analytics, Performance Acceleration,
Exit Planning, Transaction Support, Supply Chain & Operational Logistics.
Include the article's listed topics as a starting point.

TECHNOLOGY_AI: Comma-separated list of specific technologies discussed. Examples:
Machine Learning, Natural Language Processing, Robotic Process Automation,
Predictive Analytics, Cloud Computing, ERP Systems, Business Intelligence,
Generative AI, Large Language Models. Write "N/A" if no specific technology discussed.

KEYWORDS_TAGS: 5-10 comma-separated keyword tags. Include topic names, practice names,
and key themes.

BD_EMAIL_LANGUAGE: Write a 3-4 sentence email paragraph for business development
professionals to share this content with clients. Professional, consultative tone.
No greeting or sign-off. Reference the key insight and its relevance."""

    user_prompt = f"""TITLE: {title}
TYPE: {content_type}
TOPICS: {topics}
AUTHOR: {author_str}

ARTICLE TEXT:
{body_text[:8000]}"""  # Truncate to fit token limits

    try:
        response = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.3,
            max_tokens=2000,
        )
        result = json.loads(response.choices[0].message.content)
        return result
    except Exception as e:
        logging.error(f"AI enrichment failed: {e}")
        # Return fallback values
        return {
            "summary": body_text[:200] + "..." if len(body_text) > 200 else body_text,
            "qa": "",
            "audience": "PE Partners, Portfolio Company CFOs",
            "industry": "Private Equity",
            "geography": "Global",
            "solutions_value_creation": topics,
            "technology_ai": "N/A",
            "keywords_tags": topics,
            "bd_email_language": f"I wanted to share a recent {content_type.lower()} from our team: {title}.",
        }


# ---------------------------------------------------------------------------
# SharePoint Excel writer
# ---------------------------------------------------------------------------
def write_row_to_excel(token: str, row_values: list) -> None:
    """
    Add a new row to the Excel table in SharePoint using Microsoft Graph.
    row_values is a list of 17 strings, one per column.
    """
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}"
        f"/drive/items/{SHAREPOINT_FILE_ID}"
        f"/workbook/tables/{SHAREPOINT_TABLE_NAME}/rows/add"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    body = {"values": [row_values]}
    resp = requests.post(url, headers=headers, json=body)
    resp.raise_for_status()
    logging.info(f"Row written successfully: {row_values[0]}")


# ---------------------------------------------------------------------------
# Main orchestrator
# ---------------------------------------------------------------------------
def process_new_articles() -> dict:
    """
    Main function: fetch new articles, enrich with AI, write to SharePoint.
    Returns a summary of what was processed.
    """
    logging.info("Starting article processing run...")

    # Step 1: Authenticate with Microsoft Graph
    graph_token = get_graph_token()

    # Step 2: Get last processed ID
    last_id = get_last_processed_id(graph_token)
    logging.info(f"Last processed article ID: {last_id}")

    # Step 3: Fetch recent articles from WordPress
    articles = fetch_recent_articles(count=10)
    logging.info(f"Fetched {len(articles)} articles from WordPress API")

    # Step 4: Identify new articles (ID > last_id)
    new_articles = [a for a in articles if a["id"] > last_id]
    # Sort oldest first so we process in chronological order
    new_articles.sort(key=lambda a: a["id"])
    logging.info(f"Found {len(new_articles)} new articles to process")

    if not new_articles:
        return {"status": "no_new_articles", "count": 0}

    processed = []
    errors = []
    highest_id = last_id

    for article in new_articles:
        article_id = article["id"]
        title = html_module.unescape(article["title"]["rendered"])
        logging.info(f"Processing article {article_id}: {title}")

        try:
            # Step 5: Fetch full article with embedded data
            full_article = fetch_article_detail(article_id)

            # Step 6: Parse HTML content
            parsed = parse_article_html(full_article)

            # Step 7: Map taxonomy fields
            content_type = map_knowledge_type(article.get("knowledge_type", []))
            topics = map_topics(article.get("topics", []))

            # Format authors
            author_str = "; ".join(
                f"{a['name']}, {a['title']}" if a['title'] else a['name']
                for a in parsed["authors"]
            ) if parsed["authors"] else ""

            # Step 8: Generate AI enrichment
            ai_data = generate_ai_enrichment(
                title=title,
                body_text=parsed["body_text"],
                content_type=content_type,
                topics=topics,
                author_str=author_str,
            )

            # Step 9: Format the publish date
            publish_date = datetime.fromisoformat(
                article["date"]
            ).strftime("%Y-%m-%d")

            # Step 10: Assemble the 17-column row
            row = [
                title,                                              # 1. Topic Title
                content_type,                                       # 2. Type
                ai_data.get("summary", ""),                         # 3. Summary
                ai_data.get("qa", ""),                              # 4. Q&A
                author_str,                                         # 5. Authors
                publish_date,                                       # 6. Publish Date
                parsed.get("external_publication") or "",           # 7. Publication
                article.get("link", ""),                            # 8. URL
                "; ".join(parsed.get("pdf_links", [])),             # 9. Link to PDF
                parsed.get("external_publication_url") or "",       # 10. Publication URL
                ai_data.get("audience", ""),                        # 11. Audience
                ai_data.get("industry", ""),                        # 12. Industry
                ai_data.get("geography", ""),                       # 13. Geography
                ai_data.get("solutions_value_creation", ""),        # 14. Solutions
                ai_data.get("technology_ai", ""),                   # 15. Technology/AI
                ai_data.get("keywords_tags", ""),                   # 16. Keywords/Tags
                ai_data.get("bd_email_language", ""),               # 17. BD Email
            ]

            # Step 11: Write to SharePoint Excel
            write_row_to_excel(graph_token, row)

            processed.append({"id": article_id, "title": title})
            if article_id > highest_id:
                highest_id = article_id

        except Exception as e:
            logging.error(f"Failed to process article {article_id}: {e}")
            errors.append({"id": article_id, "title": title, "error": str(e)})

    # Step 12: Update the tracking ID
    if highest_id > last_id:
        update_last_processed_id(graph_token, highest_id)
        logging.info(f"Updated last processed ID to {highest_id}")

    return {
        "status": "completed",
        "processed_count": len(processed),
        "processed": processed,
        "error_count": len(errors),
        "errors": errors,
    }


# ---------------------------------------------------------------------------
# Azure Function HTTP trigger entry point
# ---------------------------------------------------------------------------
def main(req: func.HttpRequest) -> func.HttpResponse:
    """HTTP trigger entry point for the Azure Function."""
    logging.info("ProcessNewArticles function triggered via HTTP")

    try:
        result = process_new_articles()
        return func.HttpResponse(
            body=json.dumps(result, indent=2),
            status_code=200,
            mimetype="application/json",
        )
    except Exception as e:
        logging.error(f"Function failed: {e}")
        return func.HttpResponse(
            body=json.dumps({"status": "error", "message": str(e)}),
            status_code=500,
            mimetype="application/json",
        )
```

#### Part 2: Create the Power Automate Flow

1. **Go to Power Automate** (make.powerautomate.com).
2. **Create a new flow:** Click "Create" > "Scheduled cloud flow."
3. **Configure the trigger:**
   - Flow name: `Accordion Insights Monitor`
   - Starting: Today's date
   - Repeat every: `6` `Hours`
4. **Add an HTTP action:**
   - Click "+ New step" > Search for "HTTP"
   - Method: `POST`
   - URI: `https://func-accordion-insights.azurewebsites.net/api/ProcessNewArticles?code=YOUR_FUNCTION_KEY`
   - (You get the function key from Azure Portal > Your Function App > Functions > ProcessNewArticles > Function Keys)
5. **Add a Condition:**
   - After the HTTP action, add a Condition.
   - Value: `body('HTTP')?['processed_count']`
   - Operator: "is greater than"
   - Value: `0`
6. **If yes (new articles were processed):**
   - Add a "Post message in a chat or channel" action (Microsoft Teams).
   - Channel: Your team's channel.
   - Message: `New articles processed: @{body('HTTP')?['processed_count']}. Check the SharePoint spreadsheet.`
7. **If no:**
   - Optionally log "No new articles" or leave empty.
8. **Save and test** the flow.

---

## 7. Implementation Option B: Fully Python-Based Azure Functions

**Best for:** Developers who want full control, version control (Git), and no dependency on Power Automate.

### Architecture

```
Azure Functions App (Python 3.11+)
    |
    +--> Timer Trigger: CRON "0 0 */6 * * *"  (every 6 hours)
    |
    +--> Same Python code as Option A
    |
    +--> (No Power Automate involved)
```

### Differences from Option A

| Aspect | Option A | Option B |
|---|---|---|
| Trigger | Power Automate Recurrence | Azure Functions Timer Trigger |
| Monitoring | Power Automate run history | Azure Application Insights |
| Notifications | Power Automate Teams action | Python code sends Teams webhook |
| Deployment | Manual (Azure Portal) or VS Code | Azure Functions Core Tools CLI or GitHub Actions CI/CD |
| Version control | Azure Portal only | Full Git support |

### Timer Trigger Setup

Instead of an HTTP trigger, use a Timer trigger. The only difference is the `function.json` file:

**`function.json`:**
```json
{
  "scriptFile": "__init__.py",
  "bindings": [
    {
      "name": "timer",
      "type": "timerTrigger",
      "direction": "in",
      "schedule": "0 0 */6 * * *"
    }
  ]
}
```

**`__init__.py` (Timer version):**

```python
import azure.functions as func
import logging
import json
# Import process_new_articles from the shared module (same code as Option A)
from .core import process_new_articles


def main(timer: func.TimerRequest) -> None:
    """Timer trigger entry point -- runs every 6 hours."""
    logging.info("ProcessNewArticles triggered by timer")

    if timer.past_due:
        logging.warning("Timer is past due! Running now.")

    try:
        result = process_new_articles()
        logging.info(f"Result: {json.dumps(result)}")

        # Optionally send a Teams notification via webhook
        if result.get("processed_count", 0) > 0:
            send_teams_notification(result)

    except Exception as e:
        logging.error(f"Function failed: {e}")
        # Optionally send an alert
        send_error_alert(str(e))


def send_teams_notification(result: dict) -> None:
    """Send a summary to a Teams channel via incoming webhook."""
    import requests

    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL", "")
    if not webhook_url:
        return

    count = result["processed_count"]
    titles = [a["title"] for a in result.get("processed", [])]
    title_list = "\n".join(f"- {t}" for t in titles)

    payload = {
        "text": f"**Accordion Insights Monitor**\n\n"
                f"Processed {count} new article(s):\n{title_list}\n\n"
                f"Check the SharePoint spreadsheet for details."
    }
    requests.post(webhook_url, json=payload)


def send_error_alert(error_message: str) -> None:
    """Send an error alert to Teams."""
    import requests

    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL", "")
    if not webhook_url:
        return

    payload = {
        "text": f"**Accordion Insights Monitor - ERROR**\n\n"
                f"The automation encountered an error:\n\n`{error_message}`\n\n"
                f"Please check Azure Application Insights for details."
    }
    requests.post(webhook_url, json=payload)
```

### Project Structure (for Option B)

```
accordion-insights-automation/
|
+-- .github/
|   +-- workflows/
|       +-- deploy.yml           # GitHub Actions CI/CD pipeline
|
+-- ProcessNewArticles/
|   +-- __init__.py              # Timer trigger entry point
|   +-- function.json            # Timer trigger binding
|   +-- core.py                  # All business logic (same as Option A code)
|
+-- tests/
|   +-- test_parser.py           # Unit tests for HTML parsing
|   +-- test_taxonomy.py         # Unit tests for taxonomy mapping
|   +-- test_ai_prompt.py        # Unit tests for AI prompt formatting
|   +-- fixtures/
|       +-- sample_article.json  # Sample API response for testing
|       +-- sample_html.html     # Sample article HTML for parsing tests
|
+-- requirements.txt
+-- host.json
+-- local.settings.json          # Local development settings (DO NOT commit)
+-- .gitignore
```

**`host.json`:**
```json
{
  "version": "2.0",
  "logging": {
    "applicationInsights": {
      "samplingSettings": {
        "isEnabled": true,
        "excludedTypes": "Request"
      }
    }
  },
  "functionTimeout": "00:10:00"
}
```

**`requirements.txt`:**
```
azure-functions
requests>=2.31.0
beautifulsoup4>=4.12.0
openai>=1.12.0
msal>=1.26.0
```

**`.gitignore`:**
```
local.settings.json
.python_packages/
__pycache__/
.venv/
```

### Local Development

```bash
# Install Azure Functions Core Tools (macOS)
brew install azure-functions-core-tools@4

# Create virtual environment
python -m venv .venv
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Create local.settings.json with your environment variables
# (Copy the settings from the table in Section 6, Part 1, Step 5)

# Run locally
func start
```

### CI/CD with GitHub Actions

**`.github/workflows/deploy.yml`:**
```yaml
name: Deploy Azure Function

on:
  push:
    branches: [main]

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Run tests
        run: python -m pytest tests/

      - name: Deploy to Azure Functions
        uses: Azure/functions-action@v1
        with:
          app-name: func-accordion-insights
          package: .
          publish-profile: ${{ secrets.AZURE_FUNCTIONAPP_PUBLISH_PROFILE }}
```

---

## 8. Error Handling

This section covers every failure mode and what the automation does about it.

### 8.1 WordPress API Errors

| Error | Cause | Response |
|---|---|---|
| **HTTP 404** | The `/knowledge` endpoint changed or doesn't exist. | Log error, send Teams alert, abort run. This needs human investigation. |
| **HTTP 429** | Rate limiting (too many requests). | Wait 60 seconds and retry once. If it fails again, abort and alert. |
| **HTTP 500/502/503** | WordPress server error. | Retry up to 3 times with exponential backoff (wait 10s, 30s, 90s). If all retries fail, abort and alert. |
| **Timeout** | The API took longer than 30 seconds. | Retry once with a 60-second timeout. If it still times out, abort and alert. |
| **Malformed JSON** | The response is not valid JSON. | Log the raw response body, send alert, abort. |
| **Empty response** | The API returns `[]` (no articles). | This is normal if the site has been quiet. Log it and exit cleanly. |

**Retry logic implementation:**

```python
import time

def fetch_with_retry(url, params=None, max_retries=3, base_timeout=30):
    """
    Fetch a URL with retry logic and exponential backoff.
    """
    for attempt in range(max_retries):
        try:
            timeout = base_timeout * (attempt + 1)  # 30s, 60s, 90s
            resp = requests.get(url, params=params, timeout=timeout)

            if resp.status_code == 429:
                wait_time = int(resp.headers.get("Retry-After", 60))
                logging.warning(f"Rate limited. Waiting {wait_time}s...")
                time.sleep(wait_time)
                continue

            resp.raise_for_status()
            return resp.json()

        except requests.exceptions.Timeout:
            logging.warning(f"Timeout on attempt {attempt + 1}/{max_retries}")
            if attempt < max_retries - 1:
                time.sleep(10 * (attempt + 1))
            continue

        except requests.exceptions.RequestException as e:
            logging.error(f"Request failed on attempt {attempt + 1}: {e}")
            if attempt < max_retries - 1:
                time.sleep(10 * (attempt + 1))
            continue

    raise Exception(f"All {max_retries} retries failed for {url}")
```

### 8.2 Article Data Errors

| Error | Cause | Response |
|---|---|---|
| **Missing author** | Article doesn't have a "Meet the Author" section. | Leave the Authors column blank. This is expected for press releases and some multimedia. |
| **Missing body text** | Article has no `content.rendered`. | Skip AI enrichment. Populate only API-direct fields. Leave AI columns blank with note "Content unavailable." |
| **No PDF found** | Article doesn't have a PDF attachment. | Leave PDF column blank. This is normal for most articles. |
| **No external publication** | Article is Accordion-original. | Leave Publication and Publication URL columns blank. This is normal for most articles. |
| **HTML parsing fails** | Unexpected HTML structure. | Log the HTML snippet that caused the failure. Fall back to raw text extraction. |
| **Very long article** | Body text exceeds 8,000 characters (AI token limit). | Truncate to 8,000 characters for the AI prompt. Log a warning. The AI can still generate useful output from a truncated article. |

### 8.3 AI Generation Errors

| Error | Cause | Response |
|---|---|---|
| **HTTP 429 (rate limit)** | Too many AI requests in a short period. | Wait 60 seconds and retry. If processing multiple articles, add a 5-second delay between AI calls. |
| **HTTP 400 (bad request)** | Prompt too long or invalid. | Truncate the article text further and retry. |
| **HTTP 500 (OpenAI server error)** | Azure OpenAI service issue. | Retry up to 3 times with backoff. |
| **Invalid JSON response** | AI returned malformed JSON. | Retry once with a slightly modified prompt ("Remember to return valid JSON"). If still invalid, use fallback values. |
| **Missing fields in response** | AI omitted one or more expected keys. | Fill in missing fields with fallback values (see Column Mapping for each column's fallback). |
| **Quota exceeded** | Monthly token limit reached. | Send alert to team. Stop processing. Human needs to increase quota or wait for reset. |

**Fallback values (used when AI generation completely fails):**

```python
FALLBACK_AI_RESPONSE = {
    "summary": "",           # Will be set to first 200 chars of body text
    "qa": "",                # Left blank
    "audience": "PE Partners, Portfolio Company CFOs",
    "industry": "Private Equity",
    "geography": "Global",
    "solutions_value_creation": "",  # Will be set to article's topics
    "technology_ai": "N/A",
    "keywords_tags": "",     # Will be set to article's topics
    "bd_email_language": "", # Will be set to generic template
}
```

### 8.4 SharePoint/Graph API Errors

| Error | Cause | Response |
|---|---|---|
| **HTTP 401 (unauthorized)** | Token expired or invalid credentials. | Re-authenticate and retry once. If it fails again, the App Registration may need to be reconfigured. Send alert. |
| **HTTP 403 (forbidden)** | Insufficient permissions. | The App Registration needs `Sites.ReadWrite.All` permission with admin consent. Send alert with instructions. |
| **HTTP 404 (not found)** | The Excel file or table doesn't exist at the expected location. | Send alert. Someone may have moved or renamed the file. |
| **HTTP 409 (conflict)** | Concurrent edit conflict. | Wait 10 seconds and retry. If the file is open in Excel desktop, there could be lock issues. |
| **HTTP 429 (throttled)** | Too many Graph API calls. | Wait for the `Retry-After` header value, then retry. |
| **Row write fails** | Data format mismatch (e.g., column count mismatch). | Log the exact row data that failed. Skip this article and continue with the next one. Send alert. |

### 8.5 Alerting Strategy

All errors that require human attention should trigger a Teams notification (via webhook) with:

1. **Timestamp** of the error.
2. **Error type** (API, parsing, AI, SharePoint).
3. **Article ID and title** (if applicable).
4. **Error message** (the actual exception text).
5. **Suggested action** (e.g., "Check if the SharePoint file was moved").

Example alert message:
```
**Accordion Insights Monitor - ERROR**

Time: 2026-02-12 15:00:03 UTC
Error Type: SharePoint Write Failure
Article: "How AI Is Transforming PE Due Diligence" (ID: 48295)
Error: HTTP 404 - The specified file was not found.

Suggested Action: Verify the Excel file still exists at its expected
SharePoint location. Check if it was renamed or moved.
```

---

## 9. AI Prompt Engineering

This section contains the complete, production-ready prompt sent to Azure OpenAI. The prompt is designed to produce consistent, structured output.

### System Prompt (Full Text)

```
You are an expert content analyst for Accordion, a private equity-focused
financial consulting firm. Accordion provides operational and financial
consulting services to private equity firms and their portfolio companies.

Your job is to analyze a piece of thought leadership content (article,
white paper, event recap, press release, or multimedia summary) and
generate structured metadata for Accordion's marketing database.

You must return a valid JSON object with exactly these 9 keys. Do not
include any text outside the JSON object.

{
  "summary": "...",
  "qa": "...",
  "audience": "...",
  "industry": "...",
  "geography": "...",
  "solutions_value_creation": "...",
  "technology_ai": "...",
  "keywords_tags": "...",
  "bd_email_language": "..."
}

FIELD INSTRUCTIONS:

1. SUMMARY
   - Write a 2-3 sentence executive summary.
   - Focus on the key takeaway and who would benefit from reading it.
   - Write in third person ("This article explores..." not "We explore...").
   - Be specific about the content's value proposition.

2. QA (Question and Answer)
   - Generate 3-5 Q&A pairs based on the article content.
   - Questions should be things a PE partner or CFO would actually ask.
   - Answers should be concise (2-3 sentences each).
   - Format exactly as: "Q: [question]\nA: [answer]\n\n" for each pair.
   - If the article already contains FAQ content, incorporate and expand
     on those questions.

3. AUDIENCE
   - Comma-separated list of target reader types.
   - Choose from: PE Partners, Operating Partners, Portfolio Company CFOs,
     Portfolio Company CEOs, Finance Teams, IT Leaders, Board Members,
     Sponsors, LPs, Due Diligence Teams, Fund Administrators,
     Chief Revenue Officers, Chief Operating Officers.
   - You may add other specific roles if the content is clearly targeted
     at them (e.g., "HR Leaders" for workforce-related content).
   - List the most relevant audiences first.

4. INDUSTRY
   - Comma-separated list of industry verticals.
   - Choose from: Private Equity, Financial Services, Healthcare,
     Technology, Manufacturing, Retail & Consumer, Energy,
     Real Estate, Cross-Industry.
   - "Private Equity" should almost always be included since Accordion
     serves the PE ecosystem.
   - Use "Cross-Industry" only if the content is truly industry-agnostic.

5. GEOGRAPHY
   - Comma-separated list of geographic markets.
   - Choose from: Global, North America, United States, Europe,
     United Kingdom, Asia-Pacific, Middle East, Latin America.
   - Default to "Global" if no specific geography is mentioned.
   - Only list specific regions if the content explicitly discusses them.

6. SOLUTIONS_VALUE_CREATION
   - Comma-separated list of Accordion's service areas that apply.
   - Choose from: Foundational Accounting, FP&A Enhancement,
     Digital Finance, Data & Analytics, Performance Acceleration,
     Exit Planning, Transaction Support,
     Supply Chain & Operational Logistics, Technology Enablement.
   - The article's listed topics should be your starting point,
     but add others if the content discusses them.

7. TECHNOLOGY_AI
   - Comma-separated list of specific technologies discussed.
   - Examples: Machine Learning, Natural Language Processing,
     Robotic Process Automation, Predictive Analytics,
     Cloud Computing, ERP Systems (SAP, Oracle, NetSuite),
     Business Intelligence (Power BI, Tableau), Generative AI,
     Large Language Models, Data Lakes, APIs, Blockchain,
     Process Mining, Document Automation.
   - Write "N/A" if no specific technology is discussed.

8. KEYWORDS_TAGS
   - 5-10 comma-separated keyword tags.
   - Must include: relevant Accordion practice names, the article's
     listed topic names, and 3-5 additional thematic keywords.
   - These should be terms someone might search for to find this content.

9. BD_EMAIL_LANGUAGE
   - Write a 3-4 sentence email paragraph.
   - Tone: Professional, consultative, relationship-building.
   - Do NOT include a greeting ("Dear...") or sign-off ("Best...").
   - Reference the article's key insight.
   - Explain why it would be relevant to the recipient.
   - Include a subtle call to action (e.g., "I'd welcome the opportunity
     to discuss how these insights apply to your portfolio").
   - Do NOT mention Accordion by name more than once.
```

### User Prompt Template

```
TITLE: {title}
TYPE: {content_type}
TOPICS: {topics}
AUTHOR: {author_string}
DATE: {publish_date}

ARTICLE TEXT:
{body_text (truncated to 8000 characters)}
```

### Why These Prompt Design Choices

| Choice | Reasoning |
|---|---|
| **JSON response format** | Forces structured output that can be parsed programmatically. The `response_format: {"type": "json_object"}` parameter in the API call enforces this. |
| **Temperature 0.3** | Low temperature produces more consistent, focused output. Higher temperatures (0.7+) would add creativity but reduce reliability. |
| **Explicit field names in system prompt** | The AI needs to know exactly what keys to return. Listing them with descriptions prevents key name mismatches. |
| **Predefined option lists** | Giving the AI specific options (e.g., audience labels) produces more consistent tagging than leaving it open-ended. |
| **"Almost always include Private Equity"** | Since all Accordion content is PE-focused, this instruction prevents the AI from omitting the obvious. |
| **Body text truncation at 8,000 chars** | GPT-4o has a 128K context window, but the response quality is best when the input is focused. 8,000 characters captures the substance of even long articles while leaving room for the system prompt and response tokens. |
| **Third-person instruction for summary** | Ensures the summary reads like a catalog entry, not like the article itself. |

---

## 10. Deployment Checklist

Use this checklist to set up the automation from scratch. Each step is a discrete action.

### Prerequisites (One-Time Setup)

- [ ] **Azure subscription** -- You need an active Azure subscription. If your organization uses Microsoft 365 E3/E5, you likely already have one.
- [ ] **Azure OpenAI access** -- Apply for access at https://aka.ms/oai/access. Deployment of GPT-4o takes a few minutes after approval.
- [ ] **SharePoint site** -- Identify or create the SharePoint site where the Excel spreadsheet lives.
- [ ] **Excel spreadsheet** -- Create (or verify) the Excel file with the 17 columns. Format the data as an Excel Table (Insert > Table).

### Step-by-Step Deployment

#### Azure Active Directory (App Registration)

- [ ] 1. Go to Azure Portal > Azure Active Directory > App Registrations.
- [ ] 2. Click "New Registration."
- [ ] 3. Name: `AccordionInsightsAutomation`.
- [ ] 4. Supported account types: "Accounts in this organizational directory only."
- [ ] 5. Click Register.
- [ ] 6. Copy the **Application (client) ID** -- you will need this.
- [ ] 7. Copy the **Directory (tenant) ID** -- you will need this.
- [ ] 8. Go to API Permissions > Add a permission > Microsoft Graph > Application permissions.
- [ ] 9. Add `Sites.ReadWrite.All`.
- [ ] 10. Click "Grant admin consent" (requires admin role).
- [ ] 11. Go to Certificates & Secrets > New client secret.
- [ ] 12. Description: `automation-secret`. Expiry: 24 months.
- [ ] 13. Copy the **secret value immediately** (it won't be shown again).

#### Azure OpenAI

- [ ] 14. Go to Azure Portal > Azure OpenAI.
- [ ] 15. Create or select an Azure OpenAI resource.
- [ ] 16. Go to Model Deployments > Deploy model.
- [ ] 17. Deploy `gpt-4o` with a deployment name (e.g., `gpt-4o`).
- [ ] 18. Copy the **endpoint URL** and **API key** from the Keys and Endpoint page.

#### SharePoint Configuration

- [ ] 19. Open the SharePoint site in a browser.
- [ ] 20. Get the **Site ID**: Navigate to `https://{tenant}.sharepoint.com/sites/{sitename}/_api/site/id` and copy the GUID.
- [ ] 21. Alternatively, use the Graph Explorer (https://developer.microsoft.com/graph/graph-explorer) to call `GET https://graph.microsoft.com/v1.0/sites/{tenant}.sharepoint.com:/sites/{sitename}` and get the `id` field.
- [ ] 22. Create a SharePoint List named `AutomationState` with two text columns: `Key` and `Value`.
- [ ] 23. Get the **List ID** from the list settings URL or via Graph API.
- [ ] 24. Open the Excel file and note the **table name** (click on the table > Table Design tab > Table Name).
- [ ] 25. Get the Excel file's **Drive Item ID** using Graph Explorer: `GET https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/{path-to-file}`.

#### Azure Function Deployment

- [ ] 26. Go to Azure Portal > Function App > Create.
- [ ] 27. Configure as described in Section 6 (Python 3.11, Consumption plan).
- [ ] 28. After creation, go to Configuration > Application Settings.
- [ ] 29. Add all 11 environment variables from the table in Section 6.
- [ ] 30. Deploy the Python code (either via VS Code extension, Azure Portal editor, or CLI).
- [ ] 31. Test the function by calling its URL in a browser (for HTTP trigger) or waiting for the timer (for Timer trigger).

#### Power Automate Flow (Option A only)

- [ ] 32. Create the flow as described in Section 6, Part 2.
- [ ] 33. Test the flow manually by clicking "Test" > "Manually."
- [ ] 34. Verify that a test row appears in the SharePoint Excel spreadsheet.

#### Verification

- [ ] 35. Check the SharePoint Excel spreadsheet for the test row.
- [ ] 36. Verify all 17 columns are populated correctly.
- [ ] 37. Check that the `AutomationState` list has been updated with the latest article ID.
- [ ] 38. Wait for the next scheduled run and verify it works automatically.
- [ ] 39. Intentionally trigger an error (e.g., wrong SharePoint file ID) and verify the Teams alert fires.

### Ongoing Maintenance

| Task | Frequency | Description |
|---|---|---|
| Check run history | Weekly | Review Power Automate run history or Azure Function logs for any failures. |
| Renew client secret | Before expiry | Azure App Registration client secrets expire. Renew before the expiry date. |
| Review AI output quality | Monthly | Spot-check 5-10 rows in the spreadsheet to ensure AI-generated fields are accurate and useful. |
| Update taxonomy maps | As needed | If Accordion adds new Knowledge Types or Topics on their website, update the `KNOWLEDGE_TYPE_MAP` and `TOPICS_MAP` dictionaries. |
| Monitor API changes | Quarterly | WordPress API endpoints can change after major WordPress updates. Verify the endpoints still work. |
| Review token usage | Monthly | Check Azure OpenAI token consumption to ensure you're within budget. |

---

## Appendix A: Finding SharePoint IDs

This section explains how to find the SharePoint Site ID, Drive ID, File ID, and List ID that the automation needs.

### Using Microsoft Graph Explorer

1. Go to https://developer.microsoft.com/graph/graph-explorer.
2. Sign in with your Microsoft 365 account.
3. Run these queries:

**Get Site ID:**
```
GET https://graph.microsoft.com/v1.0/sites/{your-tenant}.sharepoint.com:/sites/{your-site-name}
```
The `id` field in the response is your Site ID. It looks like: `{hostname},{site-collection-id},{web-id}`.

**Get Drive ID (the document library):**
```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```
Find the drive named "Documents" (or wherever your Excel file is). Copy its `id`.

**Get File ID:**
```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/{path/to/your/file.xlsx}
```
Copy the `id` from the response.

**Get List ID:**
```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists?$filter=displayName eq 'AutomationState'
```
Copy the `id` from the response.

### Using SharePoint Browser

If Graph Explorer feels too technical:

1. **Site ID:** Navigate to `https://{tenant}.sharepoint.com/sites/{sitename}/_api/site/id`. The page will show an XML response with the GUID.
2. **List ID:** Go to the list > Settings (gear icon) > List settings. The URL will contain `List=%7B{list-id}%7D`.

---

## Appendix B: Estimated Costs

| Resource | Pricing Tier | Estimated Monthly Cost |
|---|---|---|
| Azure Function (Consumption) | First 1M executions free | $0 (well under free tier) |
| Azure OpenAI (GPT-4o) | ~$5/1M input tokens, ~$15/1M output tokens | ~$2-5/month (processing ~20 articles) |
| Power Automate | Included with Microsoft 365 E3/E5 | $0 |
| SharePoint | Included with Microsoft 365 | $0 |
| **Total** | | **~$2-5/month** |

The cost is dominated by Azure OpenAI usage. Each article uses roughly:
- Input: ~2,000 tokens (system prompt + article text)
- Output: ~800 tokens (the JSON response)
- At ~20 articles/month: ~56,000 total tokens/month
- Cost: approximately $0.30-0.50/month at current GPT-4o pricing.

---

## Appendix C: Security Considerations

1. **API Keys:** All API keys and secrets are stored as Azure Function Application Settings (environment variables), which are encrypted at rest. They never appear in source code.

2. **No credentials in code:** The Python code reads all secrets from `os.environ`. The `local.settings.json` file (for local development) is in `.gitignore` and never committed.

3. **App Registration permissions:** The `Sites.ReadWrite.All` permission is scoped to SharePoint sites the app has been granted access to. Consider using `Sites.Selected` for even tighter scoping (requires Graph API site permission grants).

4. **Data in transit:** All API calls use HTTPS. Data is encrypted in transit between the Azure Function, WordPress API, Azure OpenAI, and SharePoint.

5. **No PII processed:** The automation only processes publicly available article content. No personal identifiable information is handled beyond author names (which are already public on the Accordion website).

6. **Secret rotation:** Set a calendar reminder to rotate the App Registration client secret before it expires (default: 24 months).

---

## Appendix D: Troubleshooting Guide

| Symptom | Likely Cause | Fix |
|---|---|---|
| Function runs but processes 0 articles every time | `last_processed_id` is set to a very high number | Check the AutomationState list in SharePoint. Reset the Value to `0` to reprocess all articles, or to the ID of the last article you want to skip. |
| Function fails with "401 Unauthorized" on Graph API | Client secret expired or permissions not granted | Renew the secret in Azure AD > App Registrations. Ensure admin consent is granted. |
| AI returns empty or garbage output | Prompt too long (token overflow) or model deployment issue | Check the body text length. Reduce the truncation limit from 8,000 to 4,000 characters. Verify the model deployment exists in Azure OpenAI Studio. |
| Excel row has wrong number of columns | Column count mismatch between code and spreadsheet | Verify the Excel table has exactly 17 columns in the exact order listed in Section 5. |
| HTML parsing returns no author | Accordion changed their HTML template | Inspect a recent article's HTML in browser DevTools. Update the parsing selectors in `parse_article_html()`. |
| Duplicate rows in spreadsheet | `last_processed_id` not updating | Check the AutomationState list for the correct item. Verify the function has write permission to the list. |
| Power Automate flow shows "failed" | HTTP action timed out | Increase the timeout in Power Automate HTTP action settings (default is 120 seconds; set to 300 for safety). |
| Articles appear out of order | Articles processed in wrong sequence | Verify the `new_articles.sort(key=lambda a: a["id"])` line is present. Articles should be sorted oldest-first before processing. |
