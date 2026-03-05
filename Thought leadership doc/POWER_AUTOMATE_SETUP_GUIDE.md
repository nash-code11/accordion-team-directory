# Power Automate Setup Guide: Accordion Thought Leadership Automation

## What This Guide Will Help You Build

By the end of this guide, you will have a fully automated system that:

1. Checks the Accordion website every 6 hours for new thought leadership articles
2. Scrapes each new article for details (author, body text, PDFs, external links)
3. Sends the article text to AI (OpenAI) to generate summaries, Q&A, keywords, tags, and a BD email
4. Writes a new row to your SharePoint Excel spreadsheet with all 17 columns filled in

**Time required:** 2-4 hours for Option A (Power Automate only), 3-5 hours for Option B (Azure Function)

**Difficulty:** Beginner-friendly. This guide assumes you have never used Power Automate, Azure, or APIs before.

---

## Table of Contents

- [PART 1: Prerequisites (What You Need Before Starting)](#part-1-prerequisites-what-you-need-before-starting)
- [PART 2: Prepare Your Spreadsheet](#part-2-prepare-your-spreadsheet)
- [PART 3: Build the Flow -- Option A (Simple, No Azure Function)](#part-3-build-the-power-automate-flow--option-a-simple-no-azure-function)
- [PART 4: Build the Flow -- Option B (Azure Function)](#part-4-build-the-power-automate-flow--option-b-azure-function)
- [PART 5: Testing Your Flow](#part-5-testing-your-flow)
- [PART 6: Monitoring and Maintenance](#part-6-monitoring-and-maintenance)
- [PART 7: Troubleshooting](#part-7-troubleshooting)

---

## PART 1: Prerequisites (What You Need Before Starting)

Before you build anything, you need three things set up:

1. A Microsoft 365 account with Power Automate access
2. An OpenAI API key (for the AI-generated content)
3. The SharePoint spreadsheet prepared with the correct column headers

This section walks through each one.

---

### 1.1 Microsoft 365 Account with Power Automate Access

Power Automate is included with most Microsoft 365 business plans. Here is how to check if you have access:

**Step 1:** Open your web browser and go to:
```
https://flow.microsoft.com
```

**Step 2:** Sign in with your Accordion Microsoft 365 email and password (the same one you use for Outlook, Teams, and SharePoint).

**Step 3:** After signing in, you should see the Power Automate home page. It will show a left sidebar with options like "My flows," "Create," "Templates," etc.

- If you see this page, you have access. Move on to Section 1.2.
- If you get an error saying you do not have a license, contact your IT administrator and ask them to assign you a **Power Automate** license. The "Power Automate for Microsoft 365" license (included with most E3/E5 plans) is sufficient for this project.

**Important note about the HTTP connector:** The flow we are building uses the "HTTP" action to call external APIs (the WordPress API and the OpenAI API). The HTTP action requires a **Power Automate Premium** license (previously called "Power Automate per user"). If you do not have this license, you will see an error when you try to add the HTTP action. Ask your IT administrator to assign you the Premium license. It costs approximately $15/user/month. There is no way around this requirement for Option A.

**How to check your license:**
1. Go to https://flow.microsoft.com
2. Click the gear icon (Settings) in the top-right corner
3. Click "View my licenses"
4. Look for "Power Automate Premium" or "Power Automate per user" in the list

---

### 1.2 Getting an OpenAI API Key

You have two options for AI: **OpenAI** (recommended for simplicity) or **Azure OpenAI** (recommended for enterprise compliance). This guide covers both but defaults to OpenAI because setup is faster.

#### Option 1: OpenAI API Key (Recommended for Getting Started)

**Step 1:** Open your browser and go to:
```
https://platform.openai.com/signup
```

**Step 2:** Create an account. You can sign up with your email address or use Google/Microsoft sign-in. If you already have a ChatGPT account, you can use those same credentials -- but you still need to visit the platform site separately.

**Step 3:** After signing in, you will land on the OpenAI platform dashboard. Click your profile icon in the top-right corner, then click **"API keys"** from the dropdown menu. Alternatively, go directly to:
```
https://platform.openai.com/api-keys
```

**Step 4:** Click the green **"+ Create new secret key"** button.

**Step 5:** Give your key a name (for example, "Accordion Automation") and click **"Create secret key."**

**Step 6:** A popup will appear showing your API key. It looks like a long string starting with `sk-`. **Copy this key immediately and save it somewhere safe** (a password manager, a secure note, etc.). You will NOT be able to see this key again after you close the popup. If you lose it, you will need to create a new one.

**Step 7:** Set up billing. OpenAI charges per API call (per token, which is roughly per word). Go to:
```
https://platform.openai.com/account/billing
```
Click "Add payment method" and add a credit card. Then click "Set up paid account." For this automation, costs will be very low -- roughly $0.01 to $0.05 per article processed (using GPT-4o). At a few articles per week, expect to spend less than $5/month.

**Step 8:** (Optional but recommended) Set a monthly spending limit. On the billing page, click "Usage limits" and set a hard limit (for example, $20/month). This prevents surprise charges if something goes wrong.

#### Option 2: Azure OpenAI (Enterprise/Compliance)

If your organization requires data to stay within Microsoft's Azure infrastructure, use Azure OpenAI instead. Setup is more involved:

**Step 1:** You need an Azure subscription. Go to:
```
https://portal.azure.com
```
Sign in with your Microsoft 365 credentials. If your organization already has Azure, you should see the Azure portal. If not, you may need to create a subscription (see Part 4 for Azure account setup details).

**Step 2:** In the Azure portal, search for "Azure OpenAI" in the top search bar. Click on "Azure OpenAI."

**Step 3:** Click **"+ Create"** to create a new Azure OpenAI resource.
- **Subscription:** Select your Azure subscription
- **Resource group:** Create a new one called "AccordionAutomation" or select an existing one
- **Region:** Select "East US" or whichever region is closest to your team
- **Name:** Enter "accordion-openai" (or any name you prefer)
- **Pricing tier:** Select "Standard S0"

**Step 4:** Click "Review + create" then "Create." Wait for the deployment to complete (1-2 minutes).

**Step 5:** Once deployed, go to your Azure OpenAI resource. In the left sidebar, click **"Keys and Endpoint."** Copy **Key 1** and the **Endpoint URL**. Save both securely.

**Step 6:** In the left sidebar, click **"Model deployments"** then **"Manage Deployments"** (this opens Azure AI Foundry, formerly Azure OpenAI Studio). Click **"+ Create new deployment."** Select model **"gpt-4o"** and give the deployment a name (e.g., "gpt-4o"). Click "Create."

**What to save for later:**
- If using OpenAI: Your API key (the `sk-...` string)
- If using Azure OpenAI: Your Key, your Endpoint URL, and your deployment name

---

### 1.3 The SharePoint Spreadsheet

The spreadsheet where all the data will be written already exists at this URL:

```
https://accordionpartnersnyc.sharepoint.com/:x:/s/MarketingWorking/IQBshgW2zURbQ6Ir8nrPWjh0AXX5jdlH-gCmVG8VQOutSng?e=lLwgSR
```

You need to make sure it is set up correctly before building the flow. That is covered in Part 2.

---

## PART 2: Prepare Your Spreadsheet

The spreadsheet needs to have the correct column headers and be formatted as a Table. Power Automate can only write to Excel files in SharePoint if the data is in a formatted Table. This is the most important preparation step.

---

### Step 2.1: Open the Spreadsheet

**Step 1:** Open your browser and go to:
```
https://accordionpartnersnyc.sharepoint.com/:x:/s/MarketingWorking/IQBshgW2zURbQ6Ir8nrPWjh0AXX5jdlH-gCmVG8VQOutSng?e=lLwgSR
```

**Step 2:** The file will open in Excel Online (in your browser). You should see a spreadsheet. If the file is completely blank, that is fine. If it already has data, that is also fine -- we just need to make sure the headers are correct.

**Step 3:** Click **"Edit Workbook"** at the top of the page if it opens in read-only mode. Then click **"Edit in Browser"** (you do not need the desktop app).

---

### Step 2.2: Add Column Headers in Row 1

In Row 1, enter the following 18 column headers exactly as written. Column A is the tracking column; columns B through R are the 17 data columns.

| Column | Header Name |
|--------|-------------|
| A | Article ID |
| B | Topic Title |
| C | Type |
| D | Summary |
| E | Q&A |
| F | Authors |
| G | Publish Date |
| H | Publication |
| I | URL |
| J | Link to PDF |
| K | Publication URL |
| L | Audience |
| M | Industry |
| N | Geography |
| O | Solutions/Value Creation Levers |
| P | Technology/AI |
| Q | Keywords/Tags |
| R | BD Email Language |

**How to enter these:**
1. Click on cell **A1**
2. Type `Article ID` and press the **Tab** key (this moves you to B1)
3. Type `Topic Title` and press **Tab**
4. Continue for all 18 columns
5. After typing the last header ("BD Email Language") in R1, press **Enter**

**Double-check:** Make sure there are no extra spaces before or after the header names, and that spelling matches exactly.

---

### Step 2.3: Format as a Table

This step is critical. Power Automate's Excel connector can only work with formatted Tables, not plain cell ranges.

**Step 1:** Click on cell **A1** (the first header cell).

**Step 2:** Select all 18 header cells. You can do this by:
- Clicking on A1, then holding **Shift** and clicking on R1
- OR pressing **Ctrl+Shift+Right Arrow** to select across

**Step 3:** Press **Ctrl+T** on your keyboard. A dialog box will appear that says "Create Table."

**Step 4:** The dialog will show a range like `=$A$1:$R$1`. Make sure the checkbox **"My table has headers"** is checked (it should be by default).

**Step 5:** Click **OK**.

You should now see the headers change appearance -- they will have a colored background, filter dropdown arrows, and the table will be highlighted with alternating row colors. This confirms the data is now in a formatted Table.

---

### Step 2.4: Name the Table

Power Automate needs to reference this table by name. By default, Excel names it "Table1" but we want a more descriptive name.

**Step 1:** Click anywhere inside the table (on any header cell is fine).

**Step 2:** Look at the top of the Excel window. In Excel Online, when a table is selected, a **"Table"** tab appears in the ribbon at the top. Click on it.

**Step 3:** On the left side of the Table ribbon, you will see a text field labeled **"Table Name."** It currently says something like "Table1."

**Step 4:** Click on that text field, delete the current name, and type:
```
ThoughtLeadership
```

**Step 5:** Press **Enter** to confirm the name change.

**If you are using the desktop Excel app instead of Excel Online:** The table name field is in the "Table Design" tab (which appears when you click inside the table), in the "Properties" group on the far left.

---

### Step 2.5: Save the File

**Step 1:** Excel Online saves automatically to SharePoint. You should see "Saved" or "Saving..." in the top-left area near the file name.

**Step 2:** Wait a moment to make sure it says "Saved."

**Step 3:** You can now close this browser tab. The spreadsheet is ready.

---

### Step 2.6: Note the File Location for Power Automate

Power Automate needs to know where this file lives in SharePoint. Write down or remember:

- **SharePoint Site:** Marketing Working (this is the site name in the URL: `/s/MarketingWorking/`)
- **File location:** The file is in the root of the site's document library, or in a specific folder. You will browse to it when configuring the flow.

---

## PART 3: Build the Power Automate Flow -- Option A (Simple, No Azure Function)

This approach builds everything directly inside Power Automate using HTTP connectors and built-in actions. No coding required, no Azure account needed (beyond your Microsoft 365 account).

**What this flow does:**
1. Runs every 6 hours
2. Calls the WordPress REST API to get the latest 10 articles
3. Checks which articles are already in the spreadsheet
4. For each new article, fetches the full article page and extracts data
5. Calls the OpenAI API to generate AI content
6. Writes a new row to the SharePoint Excel spreadsheet

---

### Step 3.1: Create a New Scheduled Cloud Flow

**Step 1:** Open your browser and go to:
```
https://flow.microsoft.com
```

**Step 2:** Sign in with your Microsoft 365 account if prompted.

**Step 3:** In the left sidebar, click **"+ Create"** (it has a plus icon).

**Step 4:** You will see several options. Click **"Scheduled cloud flow."**

**Step 5:** A dialog box appears asking you to configure the schedule:

- **Flow name:** Type `Accordion Thought Leadership Automation`
- **Starting:** Leave this as the current date and time (or set it to the next hour)
- **Repeat every:** Type `6` in the number field
- **Frequency:** Select `Hour` from the dropdown

**Step 6:** Click **"Create."**

You should now see the flow editor with a single "Recurrence" trigger block at the top. This block is already configured to run every 6 hours. You can verify by clicking on it -- the fields should show "Interval: 6" and "Frequency: Hour."

**What you should see:** A canvas (white area) with one blue block labeled "Recurrence" with a clock icon.

---

### Step 3.2: Add HTTP Action to Call the WordPress API

Now we add an action that calls the Accordion website's API to get the latest articles.

**Step 1:** Below the Recurrence trigger, click the **"+ New step"** button (or the **"+"** icon between steps).

**Step 2:** In the search box that appears, type `HTTP` and wait for results.

**Step 3:** Under "Actions," click on **"HTTP"** (it has a globe icon). Do NOT select "HTTP + Swagger" or "HTTP Webhook" -- select the plain **"HTTP"** action.

**Note:** If you see a message about Premium connectors, this is where you need the Power Automate Premium license. If you cannot add this action, see Section 1.1 about licensing.

**Step 4:** Configure the HTTP action with these exact values:

| Field | Value |
|-------|-------|
| **Method** | `GET` (select from dropdown) |
| **URI** | `https://www.accordion.com/wp-json/wp/v2/knowledge?per_page=10&orderby=date&order=desc` |
| **Headers** | (Leave blank -- no special headers needed) |
| **Body** | (Leave blank -- GET requests do not have a body) |

**Step 5:** Click on the three dots (**...**) in the top-right corner of this action block and select **"Rename."** Rename it to:
```
Get Latest Articles from WordPress
```

**Step 6:** Click somewhere outside the action to deselect it. The action block should now show the name "Get Latest Articles from WordPress."

---

### Step 3.3: Parse the WordPress API Response

The HTTP action returns raw data. We need to parse it so Power Automate understands the structure and we can reference individual fields later.

**Step 1:** Click **"+ New step"** below the HTTP action.

**Step 2:** Search for `Parse JSON` and click on the **"Parse JSON"** action (under "Data Operations").

**Step 3:** In the **"Content"** field, click inside it, and a panel will appear on the right showing "Dynamic content." Click on **"Body"** from the "Get Latest Articles from WordPress" section. This tells Parse JSON to parse the response body from the previous HTTP call.

**Step 4:** In the **"Schema"** field, paste the following JSON schema. This schema tells Power Automate what fields to expect in the API response:

```json
{
  "type": "array",
  "items": {
    "type": "object",
    "properties": {
      "id": {
        "type": "integer"
      },
      "date": {
        "type": "string"
      },
      "slug": {
        "type": "string"
      },
      "link": {
        "type": "string"
      },
      "title": {
        "type": "object",
        "properties": {
          "rendered": {
            "type": "string"
          }
        }
      },
      "content": {
        "type": "object",
        "properties": {
          "rendered": {
            "type": "string"
          }
        }
      },
      "knowledge_type": {
        "type": "array",
        "items": {
          "type": "integer"
        }
      },
      "topics": {
        "type": "array",
        "items": {
          "type": "integer"
        }
      }
    }
  }
}
```

**Step 5:** Rename this action to:
```
Parse WordPress Response
```

**How to use "Generate from sample" (alternative to pasting the schema):** Instead of pasting the schema above, you can click **"Generate from sample"** and paste an actual API response. To get a sample response:
1. Open a new browser tab
2. Go to `https://www.accordion.com/wp-json/wp/v2/knowledge?per_page=1&orderby=date&order=desc`
3. Your browser will show JSON data -- copy all of it
4. Paste it into the "Generate from sample" text box and click "Done"
5. Power Automate will auto-generate the schema

---

### Step 3.4: Get Existing Article IDs from the Spreadsheet

Before processing articles, we need to know which ones are already in the spreadsheet so we do not duplicate them.

**Step 1:** Click **"+ New step."**

**Step 2:** Search for `Excel` and click on **"List rows present in a table"** (under "Excel Online (Business)").

**Step 3:** Configure the action:

| Field | Value |
|-------|-------|
| **Location** | Select `SharePoint` from the dropdown |
| **Document Library** | Select `Documents` (or the library where the file lives) |
| **File** | Browse to and select the Excel file. Click the folder icon and navigate through the SharePoint site structure to find it. The site is "Marketing Working." |
| **Table** | Select `ThoughtLeadership` from the dropdown (this is the table name you set in Part 2) |

**Note:** If you do not see the file or table, make sure you completed all of Part 2 (especially the Table formatting step). Power Automate can only see formatted Tables, not plain data ranges.

**Step 4:** Rename this action to:
```
Get Existing Articles from Spreadsheet
```

---

### Step 3.5: Build an Array of Known Article IDs

We need to extract just the Article ID values from the spreadsheet rows so we can check them against the API results.

**Step 1:** Click **"+ New step."**

**Step 2:** Search for `Select` and click on the **"Select"** action (under "Data Operations").

**Step 3:** Configure:

| Field | Value |
|-------|-------|
| **From** | Click in the field, then select `value` from "Get Existing Articles from Spreadsheet" in the Dynamic content panel |
| **Map** | Switch to "Map" mode by clicking the small toggle icon on the right side of the Map field. In the **key** field, type `id`. In the **value** field, click and select `Article ID` from "Get Existing Articles from Spreadsheet" in Dynamic content. |

**Step 4:** Rename this action to:
```
Extract Known Article IDs
```

**Step 5:** Now we need to convert this to a simple text string we can search. Click **"+ New step,"** search for `Compose` (under "Data Operations"), and click it.

**Step 6:** In the **Inputs** field of the Compose action, type the following expression. To enter an expression, click on the field, then click the **"Expression"** tab (next to "Dynamic content") in the panel that appears on the right. Type or paste this expression and click **OK**:

```
join(body('Extract_Known_Article_IDs'), ',')
```

**Step 7:** Rename this Compose action to:
```
Known IDs as Text
```

---

### Step 3.6: Loop Through Each Article from the API

Now we loop through each article returned by the WordPress API and check if it is new.

**Step 1:** Click **"+ New step."**

**Step 2:** Search for `Apply to each` and click on the **"Apply to each"** control (under "Control").

**Step 3:** In the **"Select an output from previous steps"** field, click and select `body` from "Parse WordPress Response" in Dynamic content. This is the parsed array of articles.

**Step 4:** Rename this action to:
```
Process Each Article
```

Everything from this point forward goes INSIDE this "Apply to each" loop.

---

### Step 3.7: Check If the Article Is New (Condition)

Inside the loop, we check whether the current article's ID already exists in our spreadsheet.

**Step 1:** Inside the "Process Each Article" loop, click **"Add an action."**

**Step 2:** Search for `Condition` and click on the **"Condition"** control.

**Step 3:** Configure the condition:

- **Left side:** Click in the field, go to the **Expression** tab, and type:
  ```
  contains(outputs('Known_IDs_as_Text'), string(items('Process_Each_Article')?['id']))
  ```
  Click **OK**.

- **Operator:** Select `is equal to` from the dropdown.

- **Right side:** Type `false` (without quotes).

**What this does:** The `contains()` function checks if the article's ID appears in our comma-separated list of known IDs. If it does NOT contain the ID (equals false), then the article is new and we should process it.

**Step 4:** Rename this Condition to:
```
Is This Article New?
```

---

### Step 3.8: Extract Article Data (Inside the "If yes" Branch)

The Condition block creates two branches: **If yes** (the article IS new) and **If no** (the article already exists). We put all our processing actions inside the **"If yes"** branch. Leave the **"If no"** branch empty.

#### 3.8.1: Decode the Article Title

WordPress returns titles with HTML entities (e.g., `&#8217;` instead of an apostrophe). We need to clean this up.

**Step 1:** Inside the **"If yes"** branch, click **"Add an action."**

**Step 2:** Search for `Compose` and add a **Compose** action.

**Step 3:** In the **Inputs** field, go to the **Expression** tab and enter:
```
decodeUriComponent(replace(replace(replace(replace(replace(items('Process_Each_Article')?['title']?['rendered'], '&#8217;', ''''), '&#8216;', ''''), '&#8211;', '-'), '&#8212;', '-'), '&amp;', '&'))
```
Click **OK**.

**Step 4:** Rename to:
```
Clean Title
```

#### 3.8.2: Resolve Knowledge Type (Article Type)

We need to map the numeric taxonomy IDs to readable names. The WordPress API returns `knowledge_type: [44]` and we need to convert that to "Articles."

**Step 1:** Add another **Compose** action inside "If yes."

**Step 2:** In the **Inputs** field, go to the **Expression** tab and enter:
```
if(contains(string(items('Process_Each_Article')?['knowledge_type']), '44'), 'Articles', if(contains(string(items('Process_Each_Article')?['knowledge_type']), '46'), 'Event Recaps', if(contains(string(items('Process_Each_Article')?['knowledge_type']), '49'), 'Multimedia', if(contains(string(items('Process_Each_Article')?['knowledge_type']), '48'), 'Press Releases', if(contains(string(items('Process_Each_Article')?['knowledge_type']), '47'), 'White Papers', 'Unknown')))))
```
Click **OK**.

**Step 3:** Rename to:
```
Resolve Knowledge Type
```

#### 3.8.3: Resolve Topics

**Step 1:** Add another **Compose** action.

**Step 2:** In the **Inputs** field, use the Expression tab. This is a longer expression, so take care to paste it exactly:
```
join(createArray(if(contains(string(items('Process_Each_Article')?['topics']), '52'), 'Advice for Sponsors & CFOs', ''), if(contains(string(items('Process_Each_Article')?['topics']), '69'), 'Artificial Intelligence', ''), if(contains(string(items('Process_Each_Article')?['topics']), '54'), 'Data & Analytics', ''), if(contains(string(items('Process_Each_Article')?['topics']), '59'), 'Digital Finance', ''), if(contains(string(items('Process_Each_Article')?['topics']), '63'), 'Exit Planning and Transaction Support', ''), if(contains(string(items('Process_Each_Article')?['topics']), '57'), 'Foundational Accounting and FP&A Enhancement', ''), if(contains(string(items('Process_Each_Article')?['topics']), '67'), 'Healthcare', ''), if(contains(string(items('Process_Each_Article')?['topics']), '58'), 'Performance Acceleration', ''), if(contains(string(items('Process_Each_Article')?['topics']), '68'), 'Supply Chain & Operational Logistics', ''), if(contains(string(items('Process_Each_Article')?['topics']), '66'), 'Tech Tutorials', '')), ', ')
```
Click **OK**.

**Step 3:** Rename to:
```
Resolve Topics
```

**Note:** This expression will include empty strings in the joined result, producing something like `"Artificial Intelligence, , Data & Analytics, , , , , , , "`. To clean this up, wrap the whole thing in a replace: We will add a cleanup step next.

#### 3.8.4: Clean Up Topics String

**Step 1:** Add another **Compose** action.

**Step 2:** In the Expression tab, enter:
```
replace(replace(replace(outputs('Resolve_Topics'), ', , ', ', '), ', , ', ', '), ', , ', ', ')
```
Then wrap it once more to trim leading/trailing commas:
```
trim(replace(replace(replace(replace(outputs('Resolve_Topics'), ', , ', ', '), ', , ', ', '), ', , ', ', '), ', , ', ', '))
```
Click **OK**.

**Step 3:** Rename to:
```
Clean Topics
```

#### 3.8.5: Strip HTML from Article Body

The article body from the API is HTML. We need plain text for the AI prompt and for the spreadsheet.

**Step 1:** Add a **Compose** action.

**Step 2:** In the Expression tab, enter:
```
replace(replace(replace(replace(replace(replace(items('Process_Each_Article')?['content']?['rendered'], '<br>', ' '), '<br/>', ' '), '<br />', ' '), '</p>', '\n'), '</div>', '\n'), '</li>', '\n')
```
Click **OK**.

**Step 3:** Rename to:
```
Replace HTML Line Breaks
```

**Step 4:** Add another Compose action. In the Expression tab:
```
trim(replace(replace(outputs('Replace_HTML_Line_Breaks'), decodeUriComponent('%0D%0A'), ' '), decodeUriComponent('%0A'), ' '))
```
Click **OK**.

**Note:** Power Automate does not have a built-in "strip HTML" function. A more thorough approach is to use the `ConvertHtmlToText` action. Search for **"Html to text"** in the action search. If this action is available in your environment, use it instead:

**Alternative (simpler) approach:**
1. Add a new action. Search for `Html to text` and select the **"Html to text"** action (under "Content Conversion").
2. In the **Content** field, select `rendered` from the article's content (from "Parse WordPress Response" dynamic content -- it will appear as `content rendered`).
3. Rename to `Article Body as Plain Text`.

This action cleanly strips all HTML tags and gives you plain text. Use this action's output for subsequent steps.

**Step 5:** Rename the final body text compose to:
```
Article Body as Plain Text
```

#### 3.8.6: Extract the Publish Date

**Step 1:** Add a **Compose** action.

**Step 2:** In the Expression tab:
```
formatDateTime(items('Process_Each_Article')?['date'], 'MMMM d, yyyy')
```
Click **OK**.

This converts the ISO date (e.g., `2026-02-12T09:00:00`) to a readable format (e.g., `February 12, 2026`).

**Step 3:** Rename to:
```
Format Publish Date
```

---

### Step 3.9: Fetch the Full Article Page for Scraping

The WordPress API gives us the article body, but we need to scrape the actual web page for additional details like author name, PDF links, and external publication links.

**Step 1:** Inside the "If yes" branch (after the Compose actions above), add a new **HTTP** action.

**Step 2:** Configure:

| Field | Value |
|-------|-------|
| **Method** | `GET` |
| **URI** | Click in the field, go to Dynamic content, and select `link` from "Parse WordPress Response." This is the full URL to the article page. |
| **Headers** | Click "Add new parameter," check "Headers," and add: Key = `User-Agent`, Value = `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36` |

**Step 3:** Rename to:
```
Fetch Article Web Page
```

#### 3.9.1: Extract Author Name from the Page

**Step 1:** Add a **Compose** action.

**Step 2:** We look for the "Meet the Author" pattern. In the Expression tab:
```
if(
  contains(body('Fetch_Article_Web_Page'), '/team/'),
  first(split(last(split(body('Fetch_Article_Web_Page'), '/team/')), '/')),
  'Accordion'
)
```
Click **OK**.

**Step 3:** Rename to:
```
Extract Author Slug
```

**Note:** This extracts the author's URL slug (e.g., "john-smith"). To get a proper name, we would need additional parsing. A simpler approach is below.

**Simpler approach for author name:** Since the article page HTML typically contains the author name in a text format near "Meet the Author," and Power Automate's expression language is limited for HTML parsing, we will let the AI extract the author name from the body text. Skip this step and include "Extract the author name if mentioned" in the AI prompt (Step 3.10).

#### 3.9.2: Extract PDF Links

**Step 1:** Add a **Compose** action.

**Step 2:** In the Expression tab:
```
if(contains(body('Fetch_Article_Web_Page'), '.pdf'),
  first(split(last(split(body('Fetch_Article_Web_Page'), 'href="')), '"')),
  '')
```

**Important caveat:** This expression is a rough heuristic that extracts the first PDF link. If the page has multiple PDFs or the HTML structure varies, it may not work perfectly. The Azure Function approach (Option B) handles this much more reliably with BeautifulSoup.

**Step 3:** Rename to:
```
Extract PDF Link
```

#### 3.9.3: Extract External Publication Links

Similar to PDF extraction, this is a heuristic approach. We look for common publication domains.

**Step 1:** Add a **Compose** action.

**Step 2:** In the Expression tab:
```
if(or(contains(body('Fetch_Article_Web_Page'), 'forbes.com'), contains(body('Fetch_Article_Web_Page'), 'fortune.com'), contains(body('Fetch_Article_Web_Page'), 'cfodive.com'), contains(body('Fetch_Article_Web_Page'), 'wsj.com'), contains(body('Fetch_Article_Web_Page'), 'bloomberg.com')), 'External publication detected - check article page', '')
```
Click **OK**.

**Step 3:** Rename to:
```
Check External Publication
```

**Note:** Fully extracting the external URL with expressions alone is fragile. We flag that an external publication exists and the user can check manually -- or use the AI to extract it from the body text.

---

### Step 3.10: Call the OpenAI API to Generate AI Content

This is the core AI step. We send the article text to OpenAI and ask it to generate all the enrichment fields at once.

**Step 1:** Inside the "If yes" branch (after the previous steps), add a new **HTTP** action.

**Step 2:** Configure the action:

| Field | Value |
|-------|-------|
| **Method** | `POST` |
| **URI** | `https://api.openai.com/v1/chat/completions` |

**Step 3:** Click **"Show advanced options"** or **"Add new parameter"** and check **"Headers"** and **"Body."**

**Step 4:** Add the following headers (click "Add new item" for each):

| Header Key | Header Value |
|------------|-------------|
| `Content-Type` | `application/json` |
| `Authorization` | `Bearer YOUR_OPENAI_API_KEY_HERE` |

Replace `YOUR_OPENAI_API_KEY_HERE` with your actual OpenAI API key from Step 1.2. For example, if your key is `sk-abc123xyz`, the value should be `Bearer sk-abc123xyz` (note the space between "Bearer" and the key).

**If using Azure OpenAI instead,** the URI and headers are different:

| Field | Value (Azure OpenAI) |
|-------|----------------------|
| **URI** | `https://YOUR-RESOURCE-NAME.openai.azure.com/openai/deployments/YOUR-DEPLOYMENT-NAME/chat/completions?api-version=2024-02-15-preview` |
| **Header: api-key** | `YOUR_AZURE_OPENAI_KEY` |
| **Header: Content-Type** | `application/json` |

Replace `YOUR-RESOURCE-NAME`, `YOUR-DEPLOYMENT-NAME`, and `YOUR_AZURE_OPENAI_KEY` with your actual Azure OpenAI values from Section 1.2.

**Step 5:** In the **Body** field, paste the following. This is the full request body including the AI prompt.

**IMPORTANT:** Where you see dynamic content references (marked with `@{...}`), you need to click in the body field, position your cursor at that location, switch to the Dynamic content tab, and insert the correct value. Alternatively, you can type the expression directly using the Expression tab.

For simplicity, here is the approach: Paste the entire body below as text first. Then we will replace the dynamic parts.

```json
{
  "model": "gpt-4o",
  "temperature": 0.3,
  "max_tokens": 3000,
  "messages": [
    {
      "role": "system",
      "content": "You are a senior content strategist at Accordion, a private equity-focused financial consulting and advisory firm. Accordion helps PE sponsors and their portfolio companies drive value creation through operational improvements, CFO services, technology enablement, and strategic transactions.\n\nYou will be given a thought leadership article from Accordion. Generate ALL of the following in a single JSON response:\n\n1. summary: A concise 2-3 sentence executive summary. Lead with the key insight, not 'This article...'. Use language appropriate for PE sponsors and CFOs.\n\n2. qa: An array of 3-5 Q&A pairs. Questions should be what a PE sponsor or CFO would actually ask. Answers should be 2-4 sentences drawn from the article.\n\n3. authors: The author name(s) extracted from the article text (look for 'Meet the Author' section or byline). If not found, return empty string.\n\n4. audience: Array of 1-5 target audience tags. Examples: PE Sponsor, Portfolio Company CFO, Operating Partner, Finance Director, Controller, Board Member.\n\n5. industry: Array of 1-5 industry tags. Examples: Healthcare, Technology, Manufacturing, Consumer/Retail, Financial Services, Business Services, Energy.\n\n6. geography: Array of 1-3 geography tags. Examples: North America, United States, Europe, Global, Cross-border.\n\n7. solutions: Array of 1-5 solutions/value creation lever tags. Examples: Revenue Growth, Cost Optimization, Working Capital Improvement, Digital Transformation, Financial Reporting, Forecasting & Budgeting.\n\n8. technology: Array of 1-5 technology/AI tags. Examples: AI/Machine Learning, Generative AI, RPA, ERP Implementation, Cloud Migration, Data Warehousing, BI & Dashboards.\n\n9. keywords: Array of 8-15 keyword tags including Accordion practice names where relevant (CFO Services, Transaction Advisory, Technology & Digital Transformation, Data & Analytics, Performance Improvement, Portfolio Operations).\n\n10. bd_email: A professional client-facing email body (150-250 words) that an account manager could send to share this article. Open with 'Dear [Name],' and close with 'Best regards,\\n[Your Name]\\nAccordion'. Be consultative, not salesy.\n\n11. publication: If the article mentions it was originally published in an external outlet (Forbes, Fortune, CFO Dive, Bloomberg, etc.), return that outlet name. Otherwise return empty string.\n\n12. publication_url: If you can identify an external publication URL from the text, return it. Otherwise return empty string.\n\nReturn ONLY valid JSON with exactly these keys. No markdown, no code fences, no extra text."
    },
    {
      "role": "user",
      "content": "Article Title: @{outputs('Clean_Title')}\n\nArticle Type: @{outputs('Resolve_Knowledge_Type')}\n\nArticle Topics: @{outputs('Clean_Topics')}\n\nArticle Text:\n@{outputs('Article_Body_as_Plain_Text')}"
    }
  ]
}
```

**To insert the dynamic content references:** After pasting the JSON, you need to replace the four `@{...}` placeholders with actual dynamic content. Here is how:

1. Find `@{outputs('Clean_Title')}` in the body text. Delete that placeholder text. Click where it was, go to the **Expression** tab in the right panel, type `outputs('Clean_Title')`, and click OK.

2. Find `@{outputs('Resolve_Knowledge_Type')}`. Delete it. Insert the expression `outputs('Resolve_Knowledge_Type')`.

3. Find `@{outputs('Clean_Topics')}`. Delete it. Insert the expression `outputs('Clean_Topics')`.

4. Find `@{outputs('Article_Body_as_Plain_Text')}`. Delete it. Insert the expression `outputs('Article_Body_as_Plain_Text')`.

**If using the "Html to text" action (from step 3.8.5 alternative):** For the article body, use the dynamic content from that action instead: select the output of "Article Body as Plain Text" from the Dynamic content tab.

**Step 6:** Rename this action to:
```
Call OpenAI for AI Content
```

**Potential issue -- body text too long:** OpenAI's GPT-4o has a token limit. If the article text is very long (more than about 10,000 words), the API call may fail. To prevent this, you can truncate the body text. Before the HTTP call to OpenAI, add a Compose action:

Expression:
```
if(greater(length(outputs('Article_Body_as_Plain_Text')), 12000), concat(substring(outputs('Article_Body_as_Plain_Text'), 0, 12000), '\n\n[Article truncated for length]'), outputs('Article_Body_as_Plain_Text'))
```

Rename it to `Truncated Body Text` and use this output in the OpenAI request body instead of `Article_Body_as_Plain_Text`.

---

### Step 3.11: Parse the OpenAI Response

**Step 1:** Add a **Parse JSON** action after the OpenAI HTTP call (still inside the "If yes" branch).

**Step 2:** In the **Content** field, go to the Expression tab and enter:
```
json(body('Call_OpenAI_for_AI_Content')?['choices'][0]?['message']?['content'])
```
Click **OK**.

This navigates into the OpenAI response structure to extract just the AI-generated JSON content.

**Step 3:** In the **Schema** field, paste:

```json
{
  "type": "object",
  "properties": {
    "summary": {
      "type": "string"
    },
    "qa": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "question": {
            "type": "string"
          },
          "answer": {
            "type": "string"
          }
        }
      }
    },
    "authors": {
      "type": "string"
    },
    "audience": {
      "type": "array",
      "items": {
        "type": "string"
      }
    },
    "industry": {
      "type": "array",
      "items": {
        "type": "string"
      }
    },
    "geography": {
      "type": "array",
      "items": {
        "type": "string"
      }
    },
    "solutions": {
      "type": "array",
      "items": {
        "type": "string"
      }
    },
    "technology": {
      "type": "array",
      "items": {
        "type": "string"
      }
    },
    "keywords": {
      "type": "array",
      "items": {
        "type": "string"
      }
    },
    "bd_email": {
      "type": "string"
    },
    "publication": {
      "type": "string"
    },
    "publication_url": {
      "type": "string"
    }
  }
}
```

**Step 4:** Rename to:
```
Parse AI Response
```

---

### Step 3.12: Format the Q&A for the Spreadsheet

The Q&A comes back as a JSON array. We need to format it as readable text for the spreadsheet cell.

**Step 1:** Add a **Compose** action.

**Step 2:** In the Expression tab:
```
join(xpath(xml(json(concat('{"root":{"item":', string(body('Parse_AI_Response')?['qa']), '}}'))), '/root/item/concat(question, " | ", answer)'), '\n\n')
```

**Note:** This expression is advanced and may cause issues. A simpler alternative that works reliably:

**Simpler alternative:** Just store the Q&A as a JSON string.

Expression:
```
string(body('Parse_AI_Response')?['qa'])
```

This will put the Q&A in the spreadsheet as `[{"question":"...","answer":"..."}]` format, which is perfectly usable.

**Step 3:** Rename to:
```
Format Q&A
```

---

### Step 3.13: Format Array Fields as Comma-Separated Text

Several AI response fields are arrays that need to be converted to comma-separated text for the spreadsheet.

Add a **Compose** action for each of these. For each one, go to the Expression tab and enter the expression, then rename appropriately.

**Audience:**
```
join(body('Parse_AI_Response')?['audience'], ', ')
```
Rename to: `Format Audience`

**Industry:**
```
join(body('Parse_AI_Response')?['industry'], ', ')
```
Rename to: `Format Industry`

**Geography:**
```
join(body('Parse_AI_Response')?['geography'], ', ')
```
Rename to: `Format Geography`

**Solutions:**
```
join(body('Parse_AI_Response')?['solutions'], ', ')
```
Rename to: `Format Solutions`

**Technology:**
```
join(body('Parse_AI_Response')?['technology'], ', ')
```
Rename to: `Format Technology`

**Keywords:**
```
join(body('Parse_AI_Response')?['keywords'], ', ')
```
Rename to: `Format Keywords`

---

### Step 3.14: Write a New Row to the Excel Spreadsheet

This is the final action that writes all the collected and generated data to your SharePoint spreadsheet.

**Step 1:** Add a new action. Search for `Excel` and select **"Add a row into a table"** (under "Excel Online (Business)").

**Step 2:** Configure the location fields (same as Step 3.4):

| Field | Value |
|-------|-------|
| **Location** | `SharePoint` |
| **Document Library** | `Documents` |
| **File** | Browse to and select your Excel file |
| **Table** | `ThoughtLeadership` |

**Step 3:** After selecting the table, Power Automate will load all the column headers from your spreadsheet. You will see a field for each column. Fill them in as follows:

| Column Field | What to Put In It |
|--------------|-------------------|
| **Article ID** | Expression: `items('Process_Each_Article')?['id']` |
| **Topic Title** | Dynamic content: Select the output of `Clean Title` |
| **Type** | Dynamic content: Select the output of `Resolve Knowledge Type` |
| **Summary** | Dynamic content: Select `summary` from `Parse AI Response` |
| **Q&A** | Dynamic content: Select the output of `Format Q&A` |
| **Authors** | Dynamic content: Select `authors` from `Parse AI Response` |
| **Publish Date** | Dynamic content: Select the output of `Format Publish Date` |
| **Publication** | Dynamic content: Select `publication` from `Parse AI Response` |
| **URL** | Dynamic content: Select `link` from `Parse WordPress Response` |
| **Link to PDF** | Dynamic content: Select the output of `Extract PDF Link` |
| **Publication URL** | Dynamic content: Select `publication_url` from `Parse AI Response` |
| **Audience** | Dynamic content: Select the output of `Format Audience` |
| **Industry** | Dynamic content: Select the output of `Format Industry` |
| **Geography** | Dynamic content: Select the output of `Format Geography` |
| **Solutions/Value Creation Levers** | Dynamic content: Select the output of `Format Solutions` |
| **Technology/AI** | Dynamic content: Select the output of `Format Technology` |
| **Keywords/Tags** | Dynamic content: Select the output of `Format Keywords` |
| **BD Email Language** | Dynamic content: Select `bd_email` from `Parse AI Response` |

**Step 4:** Rename this action to:
```
Write Row to Spreadsheet
```

---

### Step 3.15: Save Your Flow

**Step 1:** Click the **"Save"** button in the top-right corner of the flow editor.

**Step 2:** Wait for the "Flow saved successfully" notification.

**Step 3:** If there are any errors, Power Automate will highlight them in red. Click on the error to see what field is missing or incorrectly configured.

---

### Step 3.16: Complete Flow Summary

Here is what your completed flow should look like, from top to bottom:

```
1. Recurrence (every 6 hours)
    |
2. Get Latest Articles from WordPress (HTTP GET)
    |
3. Parse WordPress Response (Parse JSON)
    |
4. Get Existing Articles from Spreadsheet (List rows present in a table)
    |
5. Extract Known Article IDs (Select)
    |
6. Known IDs as Text (Compose)
    |
7. Process Each Article (Apply to each)
    |
    +-- 8. Is This Article New? (Condition)
        |
        +-- If yes:
        |   |
        |   +-- 9. Clean Title (Compose)
        |   +-- 10. Resolve Knowledge Type (Compose)
        |   +-- 11. Resolve Topics (Compose)
        |   +-- 12. Clean Topics (Compose)
        |   +-- 13. Article Body as Plain Text (Html to text / Compose)
        |   +-- 14. Format Publish Date (Compose)
        |   +-- 15. Fetch Article Web Page (HTTP GET)
        |   +-- 16. Extract PDF Link (Compose)
        |   +-- 17. Check External Publication (Compose)
        |   +-- 18. Truncated Body Text (Compose) [optional]
        |   +-- 19. Call OpenAI for AI Content (HTTP POST)
        |   +-- 20. Parse AI Response (Parse JSON)
        |   +-- 21. Format Q&A (Compose)
        |   +-- 22. Format Audience (Compose)
        |   +-- 23. Format Industry (Compose)
        |   +-- 24. Format Geography (Compose)
        |   +-- 25. Format Solutions (Compose)
        |   +-- 26. Format Technology (Compose)
        |   +-- 27. Format Keywords (Compose)
        |   +-- 28. Write Row to Spreadsheet (Add a row into a table)
        |
        +-- If no:
            (empty -- do nothing)
```

**Total actions:** Approximately 28 (depending on whether you use optional steps).

---

## PART 4: Build the Power Automate Flow -- Option B (Azure Function)

Option B uses an Azure Function to handle the heavy lifting (API calls, HTML scraping, AI processing) and a simple Power Automate flow just to trigger it on schedule. This is more reliable than Option A because:

- Python's BeautifulSoup library does a far better job of parsing HTML than Power Automate expressions
- Error handling is more robust
- The Python script (`accordion_scraper.py`) is already written and tested
- Easier to debug and modify

---

### Step 4.1: Create an Azure Account (Free Tier)

If you do not already have an Azure account, you can create one for free. Azure Functions has a generous free tier (1 million executions per month and 400,000 GB-seconds of compute time).

**Step 1:** Open your browser and go to:
```
https://azure.microsoft.com/en-us/free/
```

**Step 2:** Click **"Start free"** or **"Try Azure for free."**

**Step 3:** Sign in with your Microsoft 365 account (the same one you use for Outlook and Teams). This links your Azure account to your existing Microsoft identity.

**Step 4:** You will be asked to verify your identity with a phone number and a credit card. The credit card is for verification only -- you will not be charged as long as you stay within the free tier limits.

**Step 5:** Complete the sign-up process. You will be redirected to the Azure portal at `https://portal.azure.com`.

---

### Step 4.2: Create a Resource Group

A Resource Group is like a folder in Azure that holds all related resources together.

**Step 1:** In the Azure portal, click **"+ Create a resource"** (the big plus icon at the top of the left sidebar, or the button on the home page).

**Step 2:** In the search box, type `Resource group` and select it from the results.

**Step 3:** Click **"Create."**

**Step 4:** Fill in:

| Field | Value |
|-------|-------|
| **Subscription** | Select your subscription (usually "Azure subscription 1" or your company name) |
| **Resource group** | Type `accordion-automation` |
| **Region** | Select `East US` (or whichever region is closest to your team) |

**Step 5:** Click **"Review + create"** at the bottom, then click **"Create."**

---

### Step 4.3: Create an Azure Function App

**Step 1:** In the Azure portal, click **"+ Create a resource"** again.

**Step 2:** Search for `Function App` and click on it, then click **"Create."**

**Step 3:** Fill in the "Basics" tab:

| Field | Value |
|-------|-------|
| **Subscription** | Select your subscription |
| **Resource Group** | Select `accordion-automation` (the one you just created) |
| **Function App name** | Type `accordion-scraper` (this must be globally unique; if taken, try `accordion-scraper-123` or similar) |
| **Do you want to deploy code or a container image?** | Select `Code` |
| **Runtime stack** | Select `Python` |
| **Version** | Select `3.11` (or the latest available) |
| **Region** | Select `East US` (same as your resource group) |
| **Operating System** | Select `Linux` |
| **Hosting options and plans** | Select `Consumption (Serverless)` -- this is the free/pay-per-use tier |

**Step 4:** Click **"Next: Storage"** and leave the defaults (it will auto-create a storage account).

**Step 5:** Click **"Next: Networking"** and leave the defaults.

**Step 6:** Click **"Review + create"** then **"Create."** Wait for the deployment to complete (2-5 minutes).

**Step 7:** When deployment is done, click **"Go to resource"** to open your new Function App.

---

### Step 4.4: Set Up Environment Variables

The Python script needs several environment variables (API keys, SharePoint IDs, etc.). We set these in the Function App's configuration.

**Step 1:** In your Function App, click **"Configuration"** in the left sidebar (under "Settings").

**Step 2:** You will see the "Application settings" tab. Click **"+ New application setting"** for each of the following:

| Name | Value | Notes |
|------|-------|-------|
| `OPENAI_API_KEY` | Your OpenAI API key (the `sk-...` string) | Required |
| `OPENAI_MODEL` | `gpt-4o` | The model to use |
| `AZURE_TENANT_ID` | Your Microsoft 365 tenant ID | See below for how to find this |
| `AZURE_CLIENT_ID` | The App Registration client ID | See Step 4.5 |
| `AZURE_CLIENT_SECRET` | The App Registration client secret | See Step 4.5 |
| `SHAREPOINT_SITE_ID` | Your SharePoint site ID | See below for how to find this |
| `SPREADSHEET_ID` | The drive item ID of the Excel file | See below for how to find this |
| `SHEET_NAME` | `Sheet1` (or whatever your worksheet tab is named) | The worksheet name |

**If using Azure OpenAI instead of OpenAI, add these additional settings:**

| Name | Value |
|------|-------|
| `OPENAI_API_BASE` | Your Azure OpenAI endpoint URL (e.g., `https://accordion-openai.openai.azure.com/`) |
| `OPENAI_API_VERSION` | `2024-02-15-preview` |

**Step 3:** After adding all settings, click **"Save"** at the top of the Configuration page. Confirm when prompted.

**How to find your Tenant ID:**
1. Go to `https://portal.azure.com`
2. Search for "Microsoft Entra ID" (formerly Azure Active Directory) in the top search bar
3. Click on it
4. On the Overview page, you will see "Tenant ID" -- copy this value

**How to find your SharePoint Site ID:**
1. Open a new browser tab
2. Go to: `https://graph.microsoft.com/v1.0/sites/accordionpartnersnyc.sharepoint.com:/sites/MarketingWorking`
3. You may need to sign in. If this URL does not work in a browser, use the Graph Explorer tool at `https://developer.microsoft.com/en-us/graph/graph-explorer`
4. In Graph Explorer, sign in, then run the query: `https://graph.microsoft.com/v1.0/sites/accordionpartnersnyc.sharepoint.com:/sites/MarketingWorking`
5. The response will include an `id` field -- this is your Site ID. It looks like `accordionpartnersnyc.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

**How to find the Spreadsheet (Drive Item) ID:**
1. In Graph Explorer, run: `https://graph.microsoft.com/v1.0/sites/{YOUR_SITE_ID}/drive/root/children`
2. This lists all files in the document library. Find your Excel file in the response and copy its `id` field.
3. Alternatively, open the Excel file in SharePoint, look at the URL bar, and find the `sourcedoc=` parameter -- decode the value between `%7B` and `%7D` (these are URL-encoded curly braces).

---

### Step 4.5: Create an Azure App Registration (for SharePoint Access)

The Python script needs permission to read from and write to SharePoint. This is done through an App Registration.

**Step 1:** In the Azure portal, search for **"App registrations"** in the top search bar and click on it.

**Step 2:** Click **"+ New registration."**

**Step 3:** Fill in:

| Field | Value |
|-------|-------|
| **Name** | `AccordionAutomation` |
| **Supported account types** | Select "Accounts in this organizational directory only" |
| **Redirect URI** | Leave blank |

**Step 4:** Click **"Register."**

**Step 5:** You will be taken to the app's overview page. Copy the following values (you will need them for the environment variables):
- **Application (client) ID** -- this is your `AZURE_CLIENT_ID`
- **Directory (tenant) ID** -- this is your `AZURE_TENANT_ID`

**Step 6:** In the left sidebar, click **"Certificates & secrets."**

**Step 7:** Click **"+ New client secret."** Enter a description (e.g., "Accordion automation secret") and select an expiry period (24 months is fine). Click **"Add."**

**Step 8:** Copy the **"Value"** column immediately (NOT the "Secret ID"). This is your `AZURE_CLIENT_SECRET`. You will NOT be able to see this value again after you navigate away.

**Step 9:** In the left sidebar, click **"API permissions."**

**Step 10:** Click **"+ Add a permission."**

**Step 11:** Click **"Microsoft Graph."**

**Step 12:** Click **"Application permissions"** (NOT "Delegated permissions").

**Step 13:** Search for and add these permissions:
- `Sites.ReadWrite.All` (allows reading and writing to SharePoint sites)
- `Files.ReadWrite.All` (allows reading and writing to files in SharePoint)

**Step 14:** After adding both permissions, click the **"Grant admin consent for [your organization]"** button. This button requires admin privileges. If you do not have admin access, ask your IT administrator to click this button for you.

**Step 15:** Verify that both permissions show a green checkmark in the "Status" column, indicating consent has been granted.

---

### Step 4.6: Deploy the Python Script to the Azure Function

The Python script (`accordion_scraper.py`) needs to be uploaded to the Function App along with a few supporting files.

#### Method 1: Deploy via VS Code (Recommended)

**Step 1:** Install Visual Studio Code (VS Code) from https://code.visualstudio.com if you do not already have it.

**Step 2:** Install the **Azure Functions extension** in VS Code:
1. Open VS Code
2. Click the Extensions icon in the left sidebar (it looks like four squares)
3. Search for "Azure Functions"
4. Click **"Install"** on the "Azure Functions" extension by Microsoft

**Step 3:** Create a new folder on your computer called `accordion-function`.

**Step 4:** Inside that folder, create the following files:

**File 1: `function_app.py`**

Create this file with the following content. This is the entry point that Azure Functions calls:

```python
import azure.functions as func
import logging
import json
import os
import re
import html
import time
from datetime import datetime, timezone

import requests
from bs4 import BeautifulSoup

app = func.FunctionApp()

# Taxonomy mappings
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

WP_API_BASE = "https://www.accordion.com/wp-json/wp/v2/knowledge"

HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
}


@app.timer_trigger(schedule="0 0 */6 * * *", arg_name="myTimer", run_on_startup=False)
def accordion_pipeline(myTimer: func.TimerRequest) -> None:
    """Run the article ingestion pipeline every 6 hours."""
    logging.info("Accordion pipeline triggered at %s", datetime.now(timezone.utc).isoformat())

    if myTimer.past_due:
        logging.info("Timer is past due. Running now.")

    try:
        _run_pipeline()
        logging.info("Pipeline completed successfully.")
    except Exception as exc:
        logging.error("Pipeline failed: %s", exc, exc_info=True)


@app.route(route="run", auth_level=func.AuthLevel.FUNCTION)
def manual_trigger(req: func.HttpRequest) -> func.HttpResponse:
    """HTTP endpoint for manual or Power Automate triggering."""
    logging.info("Manual pipeline trigger received.")
    try:
        _run_pipeline()
        return func.HttpResponse(
            json.dumps({"status": "success", "timestamp": datetime.now(timezone.utc).isoformat()}),
            mimetype="application/json",
            status_code=200,
        )
    except Exception as exc:
        logging.error("Pipeline failed: %s", exc, exc_info=True)
        return func.HttpResponse(
            json.dumps({"status": "error", "message": str(exc)}),
            mimetype="application/json",
            status_code=500,
        )


def _run_pipeline():
    """Core pipeline logic."""
    # 1. Fetch articles from WordPress API
    resp = requests.get(
        WP_API_BASE,
        params={"per_page": 10, "orderby": "date", "order": "desc"},
        headers=HTTP_HEADERS,
        timeout=30,
    )
    resp.raise_for_status()
    articles = resp.json()
    logging.info("Fetched %d articles from WordPress.", len(articles))

    # 2. Get known article IDs from SharePoint
    known_ids = _get_known_ids()
    logging.info("Found %d known article IDs.", len(known_ids))

    # 3. Filter to new articles only
    new_articles = [a for a in articles if a.get("id") not in known_ids]
    logging.info("Found %d new articles.", len(new_articles))

    if not new_articles:
        logging.info("No new articles. Done.")
        return

    # 4. Process each new article
    for article in new_articles:
        article_id = article.get("id")
        raw_title = article.get("title", {}).get("rendered", "Untitled")
        title = html.unescape(raw_title)
        article_url = article.get("link", "")
        logging.info("Processing article %s: %s", article_id, title)

        try:
            # Resolve taxonomies
            kt_ids = article.get("knowledge_type", [])
            knowledge_type = ", ".join(KNOWLEDGE_TYPE_MAP.get(k, f"Unknown ({k})") for k in kt_ids)

            topic_ids = article.get("topics", [])
            topics = ", ".join(TOPICS_MAP.get(t, f"Unknown ({t})") for t in topic_ids)
            topic_names_list = [TOPICS_MAP.get(t, "") for t in topic_ids if t in TOPICS_MAP]

            # Scrape the article page
            detail = _scrape_article(article_url) if article_url else {}
            body_text = detail.get("body_text", "")
            author_name = detail.get("author_name", "")
            pdf_links = detail.get("pdf_links", [])
            external_links = detail.get("external_links", [])

            # Format the publish date
            pub_date_raw = article.get("date", "")
            pub_date = ""
            if pub_date_raw:
                try:
                    dt = datetime.fromisoformat(pub_date_raw)
                    pub_date = dt.strftime("%B %d, %Y")
                except ValueError:
                    pub_date = pub_date_raw

            # Call OpenAI for AI content
            ai_result = _call_openai(body_text, title, knowledge_type, topics, author_name)

            # Build the row
            row = [
                article_id,
                title,
                ai_result.get("type_override", knowledge_type),
                ai_result.get("summary", ""),
                json.dumps(ai_result.get("qa", [])),
                ai_result.get("authors", author_name or ""),
                pub_date,
                ai_result.get("publication", ""),
                article_url,
                " | ".join(pdf_links) if pdf_links else "",
                ai_result.get("publication_url", ""),
                ", ".join(ai_result.get("audience", [])),
                ", ".join(ai_result.get("industry", [])),
                ", ".join(ai_result.get("geography", [])),
                ", ".join(ai_result.get("solutions", [])),
                ", ".join(ai_result.get("technology", [])),
                ", ".join(ai_result.get("keywords", [])),
                ai_result.get("bd_email", ""),
            ]

            # Write to SharePoint
            _write_to_sharepoint(row)
            logging.info("Successfully wrote article %s to SharePoint.", article_id)

        except Exception as exc:
            logging.error("Failed to process article %s: %s", article_id, exc, exc_info=True)
            continue


def _scrape_article(url):
    """Scrape an article page for author, body text, PDFs, external links."""
    time.sleep(1)
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # Extract author
    author_name = None
    team_link = soup.find("a", href=re.compile(r"/team/[\w-]+/?$"))
    if team_link:
        author_name = team_link.get_text(strip=True)

    # Extract body text
    body_text = ""
    for selector in ["div.entry-content", "article", "main"]:
        el = soup.select_one(selector)
        if el:
            body_text = el.get_text(separator="\n", strip=True)
            break

    # Extract PDF links
    pdf_links = []
    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        if href.lower().endswith(".pdf"):
            if not href.startswith("http"):
                href = f"https://www.accordion.com{href}"
            pdf_links.append(href)

    # Extract external links
    external_links = []
    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        if href.startswith("http") and "accordion.com" not in href:
            external_links.append({"url": href, "text": anchor.get_text(strip=True)})

    return {
        "author_name": author_name,
        "body_text": body_text,
        "pdf_links": list(set(pdf_links)),
        "external_links": external_links,
    }


def _call_openai(body_text, title, knowledge_type, topics, author_name):
    """Call OpenAI API to generate AI content."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        logging.warning("No OPENAI_API_KEY set. Returning empty AI content.")
        return {}

    api_base = os.getenv("OPENAI_API_BASE")
    model = os.getenv("OPENAI_MODEL", "gpt-4o")

    truncated_text = body_text[:12000]

    system_prompt = (
        "You are a senior content strategist at Accordion, a private equity-focused "
        "financial consulting and advisory firm. Accordion helps PE sponsors and their "
        "portfolio companies drive value creation through operational improvements, "
        "CFO services, technology enablement, and strategic transactions.\n\n"
        "You will be given a thought leadership article. Generate ALL of the following "
        "in a single JSON response:\n\n"
        "1. summary: 2-3 sentence executive summary. Do NOT start with 'This article'.\n"
        "2. qa: Array of 3-5 Q&A pairs as [{\"question\":\"...\",\"answer\":\"...\"}].\n"
        "3. authors: Author name(s) if found in the text, else empty string.\n"
        "4. audience: Array of 1-5 target audience tags.\n"
        "5. industry: Array of 1-5 industry tags.\n"
        "6. geography: Array of 1-3 geography tags.\n"
        "7. solutions: Array of 1-5 solutions/value creation lever tags.\n"
        "8. technology: Array of 1-5 technology/AI tags.\n"
        "9. keywords: Array of 8-15 keywords including practice names.\n"
        "10. bd_email: Professional client-facing email body (150-250 words).\n"
        "11. publication: External outlet name if mentioned, else empty string.\n"
        "12. publication_url: External publication URL if found, else empty string.\n\n"
        "Return ONLY valid JSON. No markdown, no code fences."
    )

    user_prompt = (
        f"Article Title: {title}\n"
        f"Article Type: {knowledge_type}\n"
        f"Article Topics: {topics}\n"
        f"Known Author: {author_name or 'Unknown'}\n\n"
        f"Article Text:\n{truncated_text}"
    )

    if api_base:
        # Azure OpenAI
        api_version = os.getenv("OPENAI_API_VERSION", "2024-02-15-preview")
        url = f"{api_base}/openai/deployments/{model}/chat/completions?api-version={api_version}"
        headers = {"api-key": api_key, "Content-Type": "application/json"}
    else:
        # Standard OpenAI
        url = "https://api.openai.com/v1/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    payload = {
        "model": model,
        "temperature": 0.3,
        "max_tokens": 3000,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    }

    for attempt in range(3):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=120)
            resp.raise_for_status()
            content = resp.json()["choices"][0]["message"]["content"]
            # Clean markdown fences if present
            content = re.sub(r"^```(?:json)?\s*", "", content.strip())
            content = re.sub(r"\s*```$", "", content)
            return json.loads(content)
        except Exception as exc:
            logging.warning("OpenAI attempt %d failed: %s", attempt + 1, exc)
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))

    return {}


def _get_known_ids():
    """Read existing article IDs from SharePoint Excel."""
    token = _get_graph_token()
    if not token:
        return set()

    site_id = os.getenv("SHAREPOINT_SITE_ID")
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    sheet_name = os.getenv("SHEET_NAME", "Sheet1")

    url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{spreadsheet_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='A2:A5000')"
    )

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        values = resp.json().get("values", [])
        ids = set()
        for row in values:
            if row and row[0] is not None:
                try:
                    ids.add(int(row[0]))
                except (ValueError, TypeError):
                    pass
        return ids
    except Exception as exc:
        logging.error("Failed to read known IDs: %s", exc)
        return set()


def _write_to_sharepoint(row_data):
    """Write a row to the SharePoint Excel table."""
    token = _get_graph_token()
    if not token:
        raise RuntimeError("Cannot get Graph API token")

    site_id = os.getenv("SHAREPOINT_SITE_ID")
    spreadsheet_id = os.getenv("SPREADSHEET_ID")
    sheet_name = os.getenv("SHEET_NAME", "Sheet1")

    # Try table-based approach first
    table_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{spreadsheet_id}"
        f"/workbook/worksheets('{sheet_name}')/tables"
    )
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    try:
        table_resp = requests.get(table_url, headers=headers, timeout=30)
        table_resp.raise_for_status()
        tables = table_resp.json().get("value", [])

        if tables:
            table_id = tables[0]["id"]
            add_url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{spreadsheet_id}"
                f"/workbook/worksheets('{sheet_name}')/tables('{table_id}')/rows"
            )
            payload = {"values": [row_data]}
            resp = requests.post(add_url, headers=headers, json=payload, timeout=30)
            resp.raise_for_status()
            return
    except Exception:
        pass

    # Fallback: range-based write
    used_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{spreadsheet_id}"
        f"/workbook/worksheets('{sheet_name}')/usedRange"
    )
    used_resp = requests.get(used_url, headers=headers, timeout=30)
    used_resp.raise_for_status()
    row_count = used_resp.json().get("rowCount", 1)
    next_row = row_count + 1
    num_cols = len(row_data)

    # Convert column number to letter
    n = num_cols
    end_col = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        end_col = chr(65 + r) + end_col

    range_addr = f"A{next_row}:{end_col}{next_row}"
    patch_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{spreadsheet_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{range_addr}')"
    )
    payload = {"values": [row_data]}
    resp = requests.patch(patch_url, headers=headers, json=payload, timeout=30)
    resp.raise_for_status()


def _get_graph_token():
    """Get a Microsoft Graph API access token."""
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")

    if not all([tenant_id, client_id, client_secret]):
        logging.error("Missing Azure credentials for Graph API.")
        return None

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }

    try:
        resp = requests.post(url, data=data, timeout=30)
        resp.raise_for_status()
        return resp.json().get("access_token")
    except Exception as exc:
        logging.error("Failed to get Graph token: %s", exc)
        return None
```

**File 2: `requirements.txt`**

Create a text file with this content:

```
azure-functions
requests
beautifulsoup4
```

**File 3: `host.json`**

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
  "extensionBundle": {
    "id": "Microsoft.Azure.Functions.ExtensionBundle",
    "version": "[4.*, 5.0.0)"
  }
}
```

**File 4: `local.settings.json`** (for local testing only -- do NOT deploy this file)

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "python",
    "OPENAI_API_KEY": "your-key-here",
    "AZURE_TENANT_ID": "your-tenant-id",
    "AZURE_CLIENT_ID": "your-client-id",
    "AZURE_CLIENT_SECRET": "your-client-secret",
    "SHAREPOINT_SITE_ID": "your-site-id",
    "SPREADSHEET_ID": "your-spreadsheet-id",
    "SHEET_NAME": "Sheet1"
  }
}
```

**Step 5:** Deploy using VS Code:
1. Open the `accordion-function` folder in VS Code
2. Press `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac) to open the command palette
3. Type `Azure Functions: Deploy to Function App` and select it
4. Select your subscription
5. Select `accordion-scraper` (the Function App you created)
6. Confirm the deployment when prompted
7. Wait for the deployment to complete (1-3 minutes)

#### Method 2: Deploy via Azure Portal (No VS Code Needed)

**Step 1:** In your Function App in the Azure portal, click **"Functions"** in the left sidebar.

**Step 2:** Click **"+ Create."**

**Step 3:** Select **"Timer trigger"** as the template.

**Step 4:** Set the name to `accordion_pipeline` and the schedule to `0 0 */6 * * *` (this is a CRON expression meaning "every 6 hours at minute 0 and second 0").

**Step 5:** Click **"Create."**

**Step 6:** Click on the function name to open it. Click **"Code + Test"** in the left sidebar.

**Step 7:** You will see a code editor. Replace the default code with the content of `function_app.py` above.

**Step 8:** Click **"Save."**

**Note:** Deploying via the portal's inline editor is limited -- you cannot easily add `requirements.txt` this way. The VS Code method is recommended for a complete deployment.

---

### Step 4.7: Create a Simple Power Automate Flow to Trigger the Function

If you want Power Automate to trigger the Azure Function (in addition to or instead of the timer trigger), create a simple flow:

**Step 1:** Go to https://flow.microsoft.com and create a new **Scheduled cloud flow** (same as Step 3.1). Name it `Trigger Accordion Pipeline` and set it to every 6 hours.

**Step 2:** Add an **HTTP** action with:

| Field | Value |
|-------|-------|
| **Method** | `POST` |
| **URI** | Your Azure Function URL. Find this in the Azure portal: go to your Function App, click "Functions," click on the HTTP-triggered function (`manual_trigger`), click "Get function URL," and copy it. It looks like: `https://accordion-scraper.azurewebsites.net/api/run?code=YOUR_FUNCTION_KEY` |

**Step 3:** Save the flow. Done. The Power Automate flow is just a simple scheduler that calls your Azure Function.

---

## PART 5: Testing Your Flow

---

### 5.1: Run a Manual Test (Option A -- Power Automate Flow)

**Step 1:** Go to https://flow.microsoft.com and find your flow ("Accordion Thought Leadership Automation") in the "My flows" list.

**Step 2:** Click on the flow name to open it.

**Step 3:** In the top-right corner, click the **"Test"** button.

**Step 4:** Select **"Manually"** and click **"Test."**

**Step 5:** Click **"Run flow."**

**Step 6:** Wait. The flow will execute in real time. You will see each action turn green (success) or red (failure) as it runs. The whole flow may take 2-5 minutes depending on how many articles are processed.

**Step 7:** After the flow finishes, click on any action to expand it and see:
- **Inputs:** What data went into the action
- **Outputs:** What data came out of the action
- **Duration:** How long it took

**Step 8:** Open your SharePoint Excel spreadsheet and check if new rows were added.

---

### 5.2: Run a Manual Test (Option B -- Azure Function)

**Step 1:** In the Azure portal, go to your Function App.

**Step 2:** Click **"Functions"** in the left sidebar, then click on the HTTP-triggered function (`manual_trigger`).

**Step 3:** Click **"Code + Test"** in the left sidebar.

**Step 4:** Click **"Test/Run"** at the top.

**Step 5:** Leave the default settings and click **"Run."**

**Step 6:** Wait for the response. You should see either:
- `{"status": "success", "timestamp": "2026-02-15T..."}` -- the pipeline ran successfully
- `{"status": "error", "message": "..."}` -- something went wrong; read the error message

**Step 7:** To see detailed logs, click **"Monitor"** in the left sidebar (under the function). You will see a list of recent executions with timestamps and status. Click on any execution to see the full log output.

**Step 8:** Open your SharePoint Excel spreadsheet to verify data was written.

---

### 5.3: Check the Run History

**For Power Automate (Option A):**

**Step 1:** Go to https://flow.microsoft.com

**Step 2:** Click **"My flows"** in the left sidebar.

**Step 3:** Click on your flow name.

**Step 4:** You will see a **"28-day run history"** section showing every time the flow ran, with status (Succeeded, Failed, or Cancelled) and timestamps.

**Step 5:** Click on any run to see the details of each action in that run.

**For Azure Functions (Option B):**

**Step 1:** Go to the Azure portal and navigate to your Function App.

**Step 2:** Click **"Functions"** and then click on your function.

**Step 3:** Click **"Monitor"** to see execution history.

**Step 4:** Click **"Logs"** to see real-time streaming logs (useful during testing).

---

### 5.4: Verify the Spreadsheet

**Step 1:** Open the SharePoint Excel file:
```
https://accordionpartnersnyc.sharepoint.com/:x:/s/MarketingWorking/IQBshgW2zURbQ6Ir8nrPWjh0AXX5jdlH-gCmVG8VQOutSng?e=lLwgSR
```

**Step 2:** Check that new rows have been added below the header row.

**Step 3:** Verify each column has data:
- **Article ID** should be a number (e.g., 48295)
- **Topic Title** should be the article's headline
- **Type** should be one of: Articles, Event Recaps, Multimedia, Press Releases, White Papers
- **Summary** should be 2-3 sentences
- **Q&A** should contain question and answer pairs
- **Authors** should be a name or names
- **Publish Date** should be a readable date (e.g., February 12, 2026)
- **URL** should be a full link to accordion.com
- All other columns should have relevant content

**Step 4:** If any columns are blank or contain error text, see Part 7 (Troubleshooting).

---

## PART 6: Monitoring and Maintenance

---

### 6.1: Set Up Failure Notifications

You want to be notified if the flow fails so you can fix it before too many articles are missed.

**For Power Automate (Option A):**

**Step 1:** Open your flow in the editor.

**Step 2:** At the very end of the flow (after the "Apply to each" loop), add a new **parallel branch.** To do this, click the **"+"** button after the last action, then click **"Add a parallel branch."**

**Step 3:** Actually, a simpler approach: Power Automate has built-in failure notifications. Go to your flow's detail page (not the editor), click the three dots (**...**) in the top-right, and select **"Turn on flow failure notifications."** This will email you whenever the flow fails.

**Step 4:** To set up more detailed notifications, in the flow editor:
1. After the "Apply to each" loop, add a new step
2. Search for `Send an email` and select **"Send an email (V2)"** (Office 365 Outlook)
3. Configure the action to run only when the previous action fails:
   - Click the three dots on the "Send an email" action
   - Click **"Configure run after"**
   - Uncheck "is successful"
   - Check "has failed" and "has timed out"
   - Click "Done"
4. Set the email To, Subject, and Body to alert you about the failure

**For Azure Functions (Option B):**

**Step 1:** In the Azure portal, go to your Function App.

**Step 2:** Click **"Alerts"** in the left sidebar (under Monitoring).

**Step 3:** Click **"+ Create"** then **"Alert rule."**

**Step 4:** For the Signal, select **"Http 5xx"** (server errors). Set the threshold to greater than 0.

**Step 5:** For the Action Group, create a new action group that sends an email to your address.

**Step 6:** Name the alert rule "Accordion Pipeline Failure" and click "Create."

---

### 6.2: Check If the Flow Is Running

**Weekly check:**
1. Open your flow's run history (see Section 5.3)
2. Verify that runs are happening every 6 hours
3. Verify that the most recent runs show "Succeeded"

**Monthly check:**
1. Open the SharePoint spreadsheet
2. Verify that recent articles from accordion.com/insights are appearing
3. Compare the most recent article on the website with the most recent row in the spreadsheet

---

### 6.3: What to Do If the Accordion Website Changes

Websites change their HTML structure periodically. If the Accordion website is redesigned:

**Symptoms:**
- Author names stop appearing (or show wrong data)
- Body text is empty or contains navigation/footer text
- PDF links are not being found

**How to fix:**

**For Option A (Power Automate):**
1. Visit an article page on accordion.com in your browser
2. Right-click on the page and select "View Page Source" or "Inspect"
3. Look for the new HTML structure around the author name, article body, and PDF links
4. Update your Compose expressions in the flow to match the new HTML patterns

**For Option B (Azure Function):**
1. Open the `function_app.py` file
2. Update the CSS selectors and regex patterns in the `_scrape_article()` function
3. Test locally, then redeploy to Azure

**Common changes to watch for:**
- CSS class names on the article body container (e.g., `div.entry-content` might become `div.article-body`)
- The "Meet the Author" section structure
- How PDF download links are formatted

---

### 6.4: How to Modify the AI Prompts

If the AI-generated content is not meeting your needs, you can adjust the prompts.

**For Option A (Power Automate):**
1. Open your flow in the editor
2. Find the "Call OpenAI for AI Content" HTTP action
3. Edit the Body field, specifically the `"content"` value in the `"system"` message
4. Adjust the instructions (e.g., change "2-3 sentences" to "3-5 sentences" for longer summaries)
5. Save the flow

**For Option B (Azure Function):**
1. Open `function_app.py`
2. Find the `_call_openai()` function
3. Edit the `system_prompt` string
4. Redeploy to Azure

**Tips for prompt modification:**
- Be specific about length and format requirements
- Give examples of good output if possible
- If a field is consistently wrong, add a clarifying instruction for that field
- Keep temperature at 0.3 for consistent, factual output (higher values = more creative but less predictable)

---

## PART 7: Troubleshooting

---

### Problem 1: "HTTP 403 Forbidden" When Calling the WordPress API

**What it means:** The Accordion website is blocking your request. This can happen if the site has a firewall or WAF (Web Application Firewall) that blocks non-browser requests.

**How to fix:**

1. **Add a User-Agent header.** In your HTTP action, add a header with:
   - Key: `User-Agent`
   - Value: `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36`

2. **Check if the API is still accessible.** Open this URL in your browser:
   ```
   https://www.accordion.com/wp-json/wp/v2/knowledge?per_page=1
   ```
   If it shows JSON data, the API is working. If you get an error page, the API may have been disabled or the URL may have changed.

3. **Check if your IP is blocked.** If you are testing frequently, the website's firewall may have rate-limited your IP. Wait an hour and try again. For Azure Functions, the IP will be different from your local machine.

---

### Problem 2: "The expression is invalid" in Power Automate

**What it means:** You have a syntax error in one of the expressions (the formulas you type in the Expression tab).

**How to fix:**

1. **Check for missing quotes.** Every string value in an expression must be in single quotes: `'value'` not `"value"`.

2. **Check for missing parentheses.** Every opening `(` must have a matching `)`. Count them carefully.

3. **Check the action reference names.** When you reference another action's output, the name must match EXACTLY (including underscores for spaces). For example, if your action is named "Clean Title," the expression reference is `outputs('Clean_Title')` -- spaces become underscores.

4. **Check for special characters.** If your action name contains special characters, they may be encoded differently. Look at the action's "code view" (click the three dots, then "Peek code") to see the actual internal name.

5. **Test expressions one at a time.** If a complex expression fails, break it into smaller parts. Create multiple Compose actions, each doing one simple thing, then chain them together.

**Common expression errors and fixes:**

| Error | Cause | Fix |
|-------|-------|-----|
| `The template language expression 'xxx' is not valid` | Syntax error in the expression | Check quotes, parentheses, and function names |
| `Unable to process template language expressions in action 'xxx' inputs` | Reference to a non-existent action | Verify the action name in `outputs('...')` matches exactly |
| `The template function 'xxx' is not defined` | Misspelled function name | Check the function name (e.g., `join` not `Join`, `contains` not `Contains`) |

---

### Problem 3: "Item not found" in SharePoint / Excel Errors

**What it means:** Power Automate cannot find the Excel file, the worksheet, or the table.

**How to fix:**

1. **"The file was not found":**
   - Open the SharePoint site in your browser and verify the file still exists
   - In the flow action, re-browse to the file (click the folder icon and navigate to it again)
   - Make sure the file is not checked out by another user

2. **"Table not found":**
   - Open the Excel file and verify the table still exists and is named `ThoughtLeadership`
   - Click inside the data and check the Table Design tab for the table name
   - If the table was accidentally deleted, select the data and recreate it (Ctrl+T)

3. **"The worksheet was not found":**
   - Open the Excel file and check the worksheet tab name at the bottom
   - Make sure the name in your flow matches exactly (including spaces and capitalization)

4. **"The item could not be found":** This often means the spreadsheet ID or site ID is wrong (for Option B). Double-check your environment variables.

---

### Problem 4: AI Generating Bad or Irrelevant Content

**What it means:** The OpenAI response does not match your expectations -- summaries are too generic, keywords are wrong, the email sounds off, etc.

**How to fix:**

1. **Check if the article body text is reaching the AI.** In the flow run history (or Azure Function logs), look at the input to the OpenAI call. If the article text is empty or very short, the AI has nothing to work with. Fix the HTML-to-text conversion step.

2. **The article text is too long and getting truncated.** If only the first part of the article reaches the AI, the summary may miss key points from the end. Consider increasing the truncation limit (but watch for token limits).

3. **Improve the prompt.** Add more specific instructions. For example:
   - "The summary MUST mention the specific framework or methodology discussed, not just the topic area."
   - "For keywords, ALWAYS include at least one of these Accordion practice names: CFO Services, Transaction Advisory, Technology & Digital Transformation, Data & Analytics, Performance Improvement, Portfolio Operations."
   - "For the BD email, the tone should be that of a trusted advisor sharing a relevant insight, NOT a sales pitch."

4. **The AI returns invalid JSON.** Sometimes the model wraps the response in markdown code fences (` ```json ... ``` `). The parsing step should handle this, but if it does not:
   - For Option A: Add a Compose action before Parse JSON that strips the backticks: `replace(replace(outputs('...'), '```json', ''), '```', '')`
   - For Option B: The Python code already handles this with regex

5. **Try a different model.** If using `gpt-4o-mini` (which is cheaper), try switching to `gpt-4o` for better quality. Change the model name in the HTTP body (Option A) or the `OPENAI_MODEL` environment variable (Option B).

---

### Problem 5: Rate Limits

**OpenAI rate limits:**

**What it means:** You are sending too many API requests too quickly. OpenAI returns a `429 Too Many Requests` error.

**How to fix:**

1. **Add a delay between articles.** In the "Apply to each" loop (Option A), add a **Delay** action (search for "Delay" under "Schedule") set to 5 seconds after each iteration. This slows down processing but avoids rate limits.

2. **Reduce the number of articles per run.** Change the WordPress API `per_page` parameter from 10 to 5.

3. **Check your OpenAI usage tier.** Go to https://platform.openai.com/account/limits to see your rate limits. New accounts have lower limits that increase as you use the API more.

**SharePoint / Microsoft Graph rate limits:**

**What it means:** You are making too many calls to the SharePoint API. Microsoft Graph returns a `429` error.

**How to fix:**

1. **Add delays between SharePoint write operations.** Add a 2-second delay after each "Add a row" action.

2. **Batch your writes.** Instead of writing one row at a time, collect all rows and write them at once. This is more complex to implement but more efficient.

---

### Problem 6: Flow Stops Running / Is Turned Off

**What it means:** Power Automate may automatically disable a flow if it fails repeatedly (typically after 5 consecutive failures in a row).

**How to fix:**

1. Go to https://flow.microsoft.com and find your flow in "My flows"
2. If it shows "Turned off," fix the underlying error first (check the run history for the failure reason)
3. Then click the toggle to turn it back on
4. Run a manual test to verify it works

**Prevention:** Set up failure notifications (Section 6.1) so you know about problems before the flow gets auto-disabled.

---

### Problem 7: The Flow Works in Testing But Not on Schedule

**What it means:** Manual test runs succeed, but scheduled runs fail.

**Possible causes:**

1. **Connection expiry.** Power Automate connections to SharePoint, Excel, and email can expire. Go to "Connections" in the left sidebar of Power Automate and check if any show a warning. Re-authenticate if needed.

2. **Premium license was removed.** If your Power Automate Premium license was removed or expired, HTTP actions will fail. Check your license (see Section 1.1).

3. **The schedule is not what you think.** Click on the Recurrence trigger and verify the settings. The times shown are in UTC by default -- your flow runs at those UTC times, not your local time zone. You can set a specific time zone in the Recurrence trigger's advanced options.

---

### Quick Reference: Action Names to Search For in Power Automate

| What You Need | Search For | Connector Name |
|---------------|-----------|----------------|
| Call an external API | `HTTP` | HTTP (Premium) |
| Parse JSON data | `Parse JSON` | Data Operations |
| Loop through items | `Apply to each` | Control |
| If/else condition | `Condition` | Control |
| Format/transform data | `Compose` | Data Operations |
| Select specific fields | `Select` | Data Operations |
| Get Excel rows | `List rows present in a table` | Excel Online (Business) |
| Add an Excel row | `Add a row into a table` | Excel Online (Business) |
| Convert HTML to text | `Html to text` | Content Conversion |
| Send an email | `Send an email (V2)` | Office 365 Outlook |
| Wait/pause | `Delay` | Schedule |
| Set a variable | `Initialize variable` | Variables |

---

## Appendix A: Complete Taxonomy Reference

### Knowledge Types

| ID | Name | Description |
|----|------|-------------|
| 44 | Articles | Standard thought leadership articles and blog posts |
| 46 | Event Recaps | Summaries of conferences, webinars, and events |
| 49 | Multimedia | Videos, podcasts, and interactive content |
| 48 | Press Releases | Official company announcements |
| 47 | White Papers | In-depth research papers and guides |

### Topics

| ID | Name |
|----|------|
| 52 | Advice for Sponsors & CFOs |
| 69 | Artificial Intelligence |
| 54 | Data & Analytics |
| 59 | Digital Finance |
| 63 | Exit Planning and Transaction Support |
| 57 | Foundational Accounting and FP&A Enhancement |
| 67 | Healthcare |
| 58 | Performance Acceleration |
| 68 | Supply Chain & Operational Logistics |
| 66 | Tech Tutorials |

---

## Appendix B: SharePoint Column Layout Quick Reference

| Column | Letter | Header Name | Data Source | Example Value |
|--------|--------|-------------|-------------|---------------|
| 1 | A | Article ID | WordPress API `id` field | `48295` |
| 2 | B | Topic Title | WordPress API `title.rendered` (HTML-decoded) | `How AI Is Transforming PE Due Diligence` |
| 3 | C | Type | Mapped from `knowledge_type` taxonomy ID | `Articles` |
| 4 | D | Summary | AI-generated | `PE firms are increasingly leveraging AI-driven tools to accelerate...` |
| 5 | E | Q&A | AI-generated (JSON array of question/answer pairs) | `[{"question":"How does AI improve...","answer":"AI automates..."}]` |
| 6 | F | Authors | Scraped from article page or AI-extracted | `Sarah Chen` |
| 7 | G | Publish Date | WordPress API `date` field, reformatted | `February 12, 2026` |
| 8 | H | Publication | AI-extracted (if the article ran externally) | `Forbes` or empty |
| 9 | I | URL | WordPress API `link` field | `https://www.accordion.com/knowledge/ai-transforming-pe-due-diligence/` |
| 10 | J | Link to PDF | Scraped from article page (links ending in .pdf) | `https://www.accordion.com/wp-content/uploads/2026/02/whitepaper.pdf` or empty |
| 11 | K | Publication URL | AI-extracted (external publication URL) | `https://www.forbes.com/...` or empty |
| 12 | L | Audience | AI-generated tags | `PE Sponsor, Portfolio Company CFO, Operating Partner` |
| 13 | M | Industry | AI-generated tags | `Private Equity, Financial Services, Technology` |
| 14 | N | Geography | AI-generated tags | `North America, Global` |
| 15 | O | Solutions/Value Creation Levers | AI-generated tags | `Digital Transformation, Data & Analytics, Operational Efficiency` |
| 16 | P | Technology/AI | AI-generated tags | `AI/Machine Learning, Natural Language Processing` |
| 17 | Q | Keywords/Tags | AI-generated keywords including practice names | `AI, due diligence, private equity, Digital Finance, Data & Analytics` |
| 18 | R | BD Email Language | AI-generated client-facing email | `Dear [Name], I wanted to share a recent article from our team...` |

---

## Appendix C: Cost Estimates

### Power Automate (Option A)

| Item | Cost |
|------|------|
| Power Automate Premium license | ~$15/user/month |
| Runs (4 per day, 120 per month) | Included in license |

### Azure Functions (Option B)

| Item | Cost |
|------|------|
| Azure Functions Consumption plan | Free tier covers 1M executions/month |
| Storage account | ~$0.50/month |
| **Total Azure cost** | **< $1/month** |

### OpenAI API (Both Options)

| Item | Cost |
|------|------|
| GPT-4o input tokens (~3,000 tokens per article) | ~$0.0075 per article |
| GPT-4o output tokens (~1,500 tokens per article) | ~$0.015 per article |
| **Per article** | **~$0.02** |
| **Monthly (assuming ~8 articles/month)** | **~$0.16** |
| **Monthly (assuming ~30 articles/month)** | **~$0.60** |

---

## Appendix D: Glossary for Complete Beginners

| Term | What It Means |
|------|---------------|
| **API** | Application Programming Interface. A way for programs to talk to each other. When we "call an API," we are sending a request to a server and getting structured data back. |
| **API Key** | A password-like string that identifies you when calling an API. Keep it secret. |
| **Azure** | Microsoft's cloud computing platform. Think of it as a giant computer in the sky that you can rent by the second. |
| **Azure Function** | A small program that runs in Azure. It only runs when triggered (like an alarm clock) and you only pay for the seconds it is running. |
| **Bearer Token** | A type of API key that goes in the "Authorization" header of an HTTP request. The format is always `Bearer ` followed by the token string. |
| **Compose action** | A Power Automate action that transforms data. Think of it as a scratchpad where you can do calculations or text manipulation. |
| **Connector** | In Power Automate, a pre-built integration with a service (like SharePoint, Outlook, or Excel). |
| **CRON expression** | A shorthand for defining schedules. `0 0 */6 * * *` means "at second 0, minute 0, every 6 hours, every day, every month, every day of week." |
| **Dynamic content** | In Power Automate, data from a previous step that you can insert into a later step. Like a variable. |
| **Endpoint** | A specific URL that an API exposes for a specific purpose (e.g., "the articles endpoint" is the URL that returns articles). |
| **Expression** | In Power Automate, a formula that transforms data (similar to an Excel formula). |
| **Flow** | A Power Automate workflow. A series of triggers and actions that run automatically. |
| **Graph API** | Microsoft's API for accessing Microsoft 365 data (SharePoint, Outlook, Teams, Excel, etc.). |
| **HTTP** | Hypertext Transfer Protocol. The language that web browsers and servers use to communicate. |
| **HTTP action** | A Power Automate action that makes a web request to any URL. |
| **JSON** | JavaScript Object Notation. A standard format for structured data. Looks like: `{"key": "value"}`. |
| **Parse JSON** | The process of telling a program how to read a JSON string and extract specific fields from it. |
| **Power Automate** | Microsoft's automation platform. You build "flows" by connecting triggers and actions. |
| **Premium connector** | A Power Automate connector that requires a paid license (like the HTTP connector). |
| **REST API** | A type of API that uses standard HTTP methods (GET, POST, etc.) to access data. |
| **Schema** | A description of what a JSON object looks like -- what fields it has and what types they are. |
| **SharePoint** | Microsoft's cloud platform for storing and sharing documents. Your Excel file lives here. |
| **Taxonomy** | A classification system. WordPress uses taxonomies to categorize content (like "Articles" vs "White Papers"). |
| **Trigger** | The event that starts a Power Automate flow (e.g., a schedule, a new email, a file being created). |
| **WordPress REST API** | WordPress websites automatically expose an API that lets programs read content without visiting the website in a browser. |
