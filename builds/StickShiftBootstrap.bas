Attribute VB_Name = "StickShiftBootstrap"
' =====================================================================
'  StickShift Bootstrap -- OKF-compliant
'  One-click context initialiser.
'
'  BootstrapBundle creates the standard seed concepts in an empty (or
'  partial) context root. It is idempotent: files that already exist
'  are never overwritten.
'
'  Flow:
'    1. Require context root (same guard as ApplyStickShiftWrite).
'    2. Build seed list: (_foundation/00-operating-profile.md,
'       builds/example-build.md, skills/skill-md-authoring.md).
'    3. Filter to files that do NOT yet exist (create-if-absent).
'    4. Build a <VBA_WRITE> envelope from survivors.
'    5. Apply via StickShiftWriteApply.ApplyWriteEnvelopeText (shared path).
'    6. Regenerate indexes via StickShiftIndexGenerator.GenerateStickShiftIndexes.
'    7. Summary MsgBox.
'
'  Seed content lives in the private Seed* functions below so the
'  bulky strings stay out of the other modules.
'
'  Requires: StickShiftConfig, StickShiftWriteApply, StickShiftIndexGenerator.
' =====================================================================

Option Explicit

Public Sub BootstrapBundle()
    Dim root As String
    root = StickShiftConfig.BundleRoot()
    If root = "" Then
        MsgBox "Set your context first (top bar).", vbExclamation, "StickShift"
        Exit Sub
    End If

    ' --- 1. Build the full seed list ---
    Dim seedPaths(0 To 2) As String
    Dim seedContents(0 To 2) As String

    seedPaths(0) = "_foundation/00-operating-profile.md"
    seedContents(0) = SeedOperatingProfile()

    seedPaths(1) = "builds/example-build.md"
    seedContents(1) = SeedExampleBuild()

    seedPaths(2) = "skills/skill-md-authoring.md"
    seedContents(2) = SeedSkillMdAuthoring()

    ' --- 2. Create-if-absent filter ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Build a root path with trailing separator for ResolvePath parity.
    Dim rootFwd As String
    rootFwd = root
    If Right(rootFwd, 1) <> "\" And Right(rootFwd, 1) <> "/" Then
        rootFwd = rootFwd & "\"
    End If

    Dim keepPaths(0 To 2) As String
    Dim keepContents(0 To 2) As String
    Dim keepCount As Long: keepCount = 0

    Dim i As Long
    Dim absPath As String
    For i = 0 To 2
        Dim rel As String
        rel = Replace(seedPaths(i), "/", "\")
        absPath = rootFwd & rel
        If Not fso.FileExists(absPath) Then
            keepPaths(keepCount) = seedPaths(i)
            keepContents(keepCount) = seedContents(i)
            keepCount = keepCount + 1
        End If
    Next i

    If keepCount = 0 Then
        MsgBox "Bundle already initialized - nothing to add.", vbInformation, "StickShift"
        Exit Sub
    End If

    ' --- 3. Build <VBA_WRITE> envelope from survivors ---
    Dim envelope As String
    envelope = "<VBA_WRITE>" & vbLf

    For i = 0 To keepCount - 1
        envelope = envelope & "### FILE: " & keepPaths(i) & vbLf
        envelope = envelope & keepContents(i) & vbLf
        envelope = envelope & "### END FILE" & vbLf
    Next i

    envelope = envelope & "</VBA_WRITE>"

    ' --- 4. Apply via the shared write path ---
    Dim w As Long, s As Long
    If Not StickShiftWriteApply.ApplyWriteEnvelopeText(envelope, w, s) Then
        Exit Sub
    End If

    ' --- 5. Generate indexes (includes skills manifest) ---
    StickShiftIndexGenerator.GenerateStickShiftIndexes

    ' --- 6. Summary ---
    MsgBox "Bundle initialized: " & w & " file(s) created. Indexes generated.", _
           vbInformation, "StickShift"
End Sub


' ---------------------------------------------------------------------------
'  Seed content
' ---------------------------------------------------------------------------

Private Function SeedOperatingProfile() As String
    Dim s As String
    s = "---" & vbLf
    s = s & "type: Foundation" & vbLf
    s = s & "title: Operating Profile" & vbLf
    s = s & "description: Who I am and what this bundle is for." & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "<!-- TODO: Describe who you are, your role, and what this bundle is for." & vbLf
    s = s & "     The LLM reads this file first, every session, to orient itself to" & vbLf
    s = s & "     your context. Include: your name/role, the domain this portfolio" & vbLf
    s = s & "     covers, key constraints or priorities the assistant should know, and" & vbLf
    s = s & "     any standing preferences (tone, output format, decision style). -->"
    SeedOperatingProfile = s
End Function


Private Function SeedExampleBuild() As String
    Dim s As String
    s = "---" & vbLf
    s = s & "type: Build" & vbLf
    s = s & "title: Example Build" & vbLf
    s = s & "description: Illustrates the full field set for a Build concept." & vbLf
    s = s & "status: parked" & vbLf
    s = s & "effort: M" & vbLf
    s = s & "impact: medium" & vbLf
    s = s & "last_touched: 2026-01-01" & vbLf
    s = s & "Dependencies:" & vbLf
    s = s & "  - skills/skill-md-authoring.md" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "<!-- TODO: Replace this example with a real build idea." & vbLf
    s = s & "     Field notes:" & vbLf
    s = s & "     - status drives the builds/ index grouping:" & vbLf
    s = s & "         working | idea | parked | production | archived" & vbLf
    s = s & "     - status: parked keeps this out of the single 'working' WIP slot" & vbLf
    s = s & "       and avoids tripping the stall detector." & vbLf
    s = s & "     - Dependencies lists paths to other concepts this build relies on." & vbLf
    s = s & "     - last_touched (YYYY-MM-DD) is used by the stall detector." & vbLf
    s = s & "     - effort: XS | S | M | L | XL" & vbLf
    s = s & "     - impact: low | medium | high -->"
    SeedExampleBuild = s
End Function


Private Function SeedSkillMdAuthoring() As String
    Dim s As String
    s = "---" & vbLf
    s = s & "okf_version: ""0.1""" & vbLf
    s = s & "type: Skill" & vbLf
    s = s & "title: Skill MD Authoring" & vbLf
    s = s & "description: Guides the model through gathering requirements, designing, and authoring a valid SKILL.md file that conforms to the Agent Skills specification, including YAML frontmatter and clear instructional body content. Use this skill whenever a user wants to create, improve, or validate a skill's SKILL.md based on a natural-language description, existing project files, or an existing skill directory." & vbLf
    s = s & "tags: [skill, authoring, documentation]" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "# skill-md-authoring" & vbLf
    s = s & "" & vbLf
    s = s & "## 1. Purpose" & vbLf
    s = s & "" & vbLf
    s = s & "This skill teaches you how to design and write `SKILL.md` files that comply with the Agent Skills specification." & vbLf
    s = s & "" & vbLf
    s = s & "Use this skill to:" & vbLf
    s = s & "" & vbLf
    s = s & "- Create a new `SKILL.md` from a natural-language description of a skill." & vbLf
    s = s & "- Refine or rewrite an existing `SKILL.md` for clarity and spec compliance." & vbLf
    s = s & "- Validate a draft `SKILL.md` against the required frontmatter and structural rules." & vbLf
    s = s & "- Suggest how to split detailed content into `references/` or other files." & vbLf
    s = s & "" & vbLf
    s = s & "You generate **one complete `SKILL.md` file** per invocation, unless the user explicitly asks for multiple variants." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 2. When to use this skill" & vbLf
    s = s & "" & vbLf
    s = s & "Activate this skill when:" & vbLf
    s = s & "" & vbLf
    s = s & "- The user asks to 'create a skill', 'write SKILL.md', 'define a skill', or similar." & vbLf
    s = s & "- The user provides a description of functionality and mentions the Agent Skills format." & vbLf
    s = s & "- The user wants to check if a `SKILL.md` is valid or needs corrections." & vbLf
    s = s & "- The user wants to refactor a large or messy `SKILL.md`." & vbLf
    s = s & "" & vbLf
    s = s & "If the user only wants code, scripts, or reference docs and does **not** mention a `SKILL.md` or skill definition, this skill is likely not needed." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 3. Required inputs to gather" & vbLf
    s = s & "" & vbLf
    s = s & "Before drafting the `SKILL.md`, collect the following information in natural language. Ask concise clarification questions if details are missing and the interaction model allows it." & vbLf
    s = s & "" & vbLf
    s = s & "1. **Skill purpose and scope**" & vbLf
    s = s & "   - What the skill should do." & vbLf
    s = s & "   - Typical tasks, workflows, or problems it should help with." & vbLf
    s = s & "   - Any clear in-scope / out-of-scope boundaries." & vbLf
    s = s & "" & vbLf
    s = s & "2. **When to use it**" & vbLf
    s = s & "   - Triggers or keywords in user queries that suggest this skill is relevant." & vbLf
    s = s & "   - Situations where this skill should be preferred over other skills." & vbLf
    s = s & "" & vbLf
    s = s & "3. **Environment and constraints (optional)**" & vbLf
    s = s & "   - Target product or runtime environment, if any." & vbLf
    s = s & "   - External tools or system packages that may be required." & vbLf
    s = s & "   - Network access expectations, if relevant." & vbLf
    s = s & "" & vbLf
    s = s & "4. **Metadata and ownership (optional)**" & vbLf
    s = s & "   - Author or organization." & vbLf
    s = s & "   - Version or maturity level (e.g., `0.1`, `beta`, `1.0`)." & vbLf
    s = s & "   - Any custom metadata keys the user wants." & vbLf
    s = s & "" & vbLf
    s = s & "5. **Tool usage (optional)**" & vbLf
    s = s & "   - Whether the skill should have a predefined set of allowed tools." & vbLf
    s = s & "   - Which tools and with what parameters." & vbLf
    s = s & "" & vbLf
    s = s & "6. **Existing files or structure (optional)**" & vbLf
    s = s & "   - Whether there are scripts, reference documents, or assets that this skill should reference." & vbLf
    s = s & "   - Names and relative paths of those files (e.g., `scripts/extract.py`, `references/REFERENCE.md`)." & vbLf
    s = s & "" & vbLf
    s = s & "If the user prefers not to answer follow-up questions, infer reasonable defaults from their description and clearly document assumptions in the skill body." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 4. Output format overview" & vbLf
    s = s & "" & vbLf
    s = s & "You must always output a **single, complete `SKILL.md`** file structured as:" & vbLf
    s = s & "" & vbLf
    s = s & "1. **YAML frontmatter block**, surrounded by `---` lines at the top." & vbLf
    s = s & "2. **Markdown body**, containing instructions and guidance for agents." & vbLf
    s = s & "" & vbLf
    s = s & "Example outline (not actual output):" & vbLf
    s = s & "" & vbLf
    s = s & "```markdown" & vbLf
    s = s & "---" & vbLf
    s = s & "name: some-skill" & vbLf
    s = s & "description: ..." & vbLf
    s = s & "metadata:" & vbLf
    s = s & "  author: ..." & vbLf
    s = s & "  version: ..." & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "# some-skill" & vbLf
    s = s & "" & vbLf
    s = s & "## Purpose" & vbLf
    s = s & "..." & vbLf
    s = s & "" & vbLf
    s = s & "## How to use this skill" & vbLf
    s = s & "..." & vbLf
    s = s & "" & vbLf
    s = s & "## Examples" & vbLf
    s = s & "..." & vbLf
    s = s & "```" & vbLf
    s = s & "" & vbLf
    s = s & "Keep the total file reasonably sized (aim for under 500 lines) and keep very detailed reference material in separate files that can be referenced." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 5. Frontmatter generation rules" & vbLf
    s = s & "" & vbLf
    s = s & "### 5.1 `name` (required)" & vbLf
    s = s & "" & vbLf
    s = s & "The `name` field must:" & vbLf
    s = s & "" & vbLf
    s = s & "- Be 1-64 characters." & vbLf
    s = s & "- Use only lowercase letters (`a-z`), digits (`0-9`), and hyphens (`-`)." & vbLf
    s = s & "- Not start or end with a hyphen." & vbLf
    s = s & "- Not contain consecutive hyphens (`--`)." & vbLf
    s = s & "- Match the parent directory name (the user is responsible for naming the directory, but you must propose a compliant name)." & vbLf
    s = s & "" & vbLf
    s = s & "**Behavior:**" & vbLf
    s = s & "" & vbLf
    s = s & "1. If the user provides a valid name, use it exactly." & vbLf
    s = s & "2. If the user provides an invalid name:" & vbLf
    s = s & "   - Produce a corrected version that complies with all rules." & vbLf
    s = s & "   - Keep it recognizable (e.g., `PDF Processing` -> `pdf-processing`)." & vbLf
    s = s & "3. If no name is provided:" & vbLf
    s = s & "   - Derive a concise, descriptive name from the skill's purpose." & vbLf
    s = s & "   - Prefer 2-4 short words joined by hyphens, e.g., `log-analysis`, `task-prioritization`." & vbLf
    s = s & "" & vbLf
    s = s & "Always verify that your chosen name meets the constraints before returning the result." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "### 5.2 `description` (required)" & vbLf
    s = s & "" & vbLf
    s = s & "The `description` field must:" & vbLf
    s = s & "" & vbLf
    s = s & "- Be 1-1024 characters." & vbLf
    s = s & "- Clearly describe:" & vbLf
    s = s & "  - What the skill does." & vbLf
    s = s & "  - When it should be used." & vbLf
    s = s & "" & vbLf
    s = s & "**Guidelines:**" & vbLf
    s = s & "" & vbLf
    s = s & "- Use one or two sentences, or a short paragraph." & vbLf
    s = s & "- Mention action verbs and domain terms (e.g., 'summarizes', 'transforms', 'code', 'logs', 'PDF', 'forms')." & vbLf
    s = s & "- Avoid vague descriptions like 'Helps with X'." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "### 5.3 Optional fields" & vbLf
    s = s & "" & vbLf
    s = s & "Include optional fields only when they add value or the user requests them." & vbLf
    s = s & "" & vbLf
    s = s & "#### `license` (optional)" & vbLf
    s = s & "" & vbLf
    s = s & "- Short text naming the license or pointing to a bundled file." & vbLf
    s = s & "- Example: `license: Apache-2.0` or `license: Proprietary. See LICENSE.txt for terms`." & vbLf
    s = s & "" & vbLf
    s = s & "#### `compatibility` (optional)" & vbLf
    s = s & "" & vbLf
    s = s & "- Only include if there are specific environment requirements." & vbLf
    s = s & "- 1-500 characters describing:" & vbLf
    s = s & "  - Intended runtime or product." & vbLf
    s = s & "  - Required system packages or tools." & vbLf
    s = s & "  - Any special capabilities needed (e.g., local file access)." & vbLf
    s = s & "" & vbLf
    s = s & "#### `metadata` (optional)" & vbLf
    s = s & "" & vbLf
    s = s & "- Map from string keys to string values." & vbLf
    s = s & "- Use clear, reasonably unique keys (e.g., `author`, `version`, `category`, `org-internal-id`)." & vbLf
    s = s & "- Values should be simple strings in quotes when needed." & vbLf
    s = s & "" & vbLf
    s = s & "Example:" & vbLf
    s = s & "" & vbLf
    s = s & "```yaml" & vbLf
    s = s & "metadata:" & vbLf
    s = s & "  author: example-org" & vbLf
    s = s & "  version: ""1.0""" & vbLf
    s = s & "  category: ""document-processing""" & vbLf
    s = s & "```" & vbLf
    s = s & "" & vbLf
    s = s & "#### `allowed-tools` (optional)" & vbLf
    s = s & "" & vbLf
    s = s & "- A single string listing allowed tools, separated by spaces." & vbLf
    s = s & "- Only include if the user or runtime expects this field." & vbLf
    s = s & "" & vbLf
    s = s & "Example:" & vbLf
    s = s & "" & vbLf
    s = s & "```yaml" & vbLf
    s = s & "allowed-tools: Bash(git:*) Bash(jq:*) Read" & vbLf
    s = s & "```" & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 6. Body content: structure and style" & vbLf
    s = s & "" & vbLf
    s = s & "After the frontmatter, write clear Markdown instructions. There are no strict format rules, but the content should help agents use the skill correctly and safely." & vbLf
    s = s & "" & vbLf
    s = s & "### 6.1 Recommended sections" & vbLf
    s = s & "" & vbLf
    s = s & "You can adapt or rename these sections based on the skill's domain:" & vbLf
    s = s & "" & vbLf
    s = s & "1. **Title**" & vbLf
    s = s & "" & vbLf
    s = s & "   Use a top-level heading matching the skill name:" & vbLf
    s = s & "" & vbLf
    s = s & "   ```markdown" & vbLf
    s = s & "   # skill-name" & vbLf
    s = s & "   ```" & vbLf
    s = s & "" & vbLf
    s = s & "2. **Purpose**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Summarize what the skill is for in 2-5 concise bullet points or a short paragraph." & vbLf
    s = s & "   - Clarify boundaries: what the skill does and does not do." & vbLf
    s = s & "" & vbLf
    s = s & "3. **When to use this skill**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Describe patterns in user requests that should trigger this skill." & vbLf
    s = s & "   - Describe situations where the skill is **not** appropriate." & vbLf
    s = s & "" & vbLf
    s = s & "4. **Inputs you should look for**" & vbLf
    s = s & "" & vbLf
    s = s & "   - List the key pieces of information the agent should extract from the user or context before acting." & vbLf
    s = s & "   - Mention any required parameters, file types, or data formats." & vbLf
    s = s & "" & vbLf
    s = s & "5. **Step-by-step workflow**" & vbLf
    s = s & "" & vbLf
    s = s & "   Provide a numbered list of steps the agent should follow, for example:" & vbLf
    s = s & "" & vbLf
    s = s & "   ```markdown" & vbLf
    s = s & "   1. Confirm the user's goal and any constraints." & vbLf
    s = s & "   2. Inspect any provided files or snippets." & vbLf
    s = s & "   3. Choose the appropriate subroutine or script, if available." & vbLf
    s = s & "   4. Run or simulate the core operation." & vbLf
    s = s & "   5. Summarize the result and provide next-step suggestions." & vbLf
    s = s & "   ```" & vbLf
    s = s & "" & vbLf
    s = s & "6. **Examples**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Include a few brief input/output examples." & vbLf
    s = s & "   - Use fenced code blocks where appropriate." & vbLf
    s = s & "   - Focus on realistic, representative cases." & vbLf
    s = s & "" & vbLf
    s = s & "7. **Edge cases and limitations**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Call out ambiguous or risky situations." & vbLf
    s = s & "   - Suggest how the agent should respond (e.g., ask clarifying questions, decline actions, or provide safe alternatives)." & vbLf
    s = s & "" & vbLf
    s = s & "8. **References to other files (if any)**" & vbLf
    s = s & "" & vbLf
    s = s & "   - If there are supporting files, reference them with **relative paths from the skill root**, for example:" & vbLf
    s = s & "     - `references/REFERENCE.md`" & vbLf
    s = s & "     - `scripts/run-analysis.sh`" & vbLf
    s = s & "   - Keep references one level deep and avoid complex chains." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 7. Step-by-step authoring procedure" & vbLf
    s = s & "" & vbLf
    s = s & "When you are asked to create or revise a `SKILL.md`, follow this procedure:" & vbLf
    s = s & "" & vbLf
    s = s & "1. **Understand the skill**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Read the user's description carefully." & vbLf
    s = s & "   - Identify the primary domain (e.g., documents, code, data analysis, forms, logs)." & vbLf
    s = s & "   - Identify the main operations (e.g., summarize, transform, validate, extract)." & vbLf
    s = s & "" & vbLf
    s = s & "2. **Draft a valid `name`**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Start from any provided name; normalize it if needed:" & vbLf
    s = s & "     - Lowercase it." & vbLf
    s = s & "     - Replace spaces and invalid characters with hyphens." & vbLf
    s = s & "     - Remove leading/trailing hyphens." & vbLf
    s = s & "     - Ensure no consecutive hyphens." & vbLf
    s = s & "   - If no name is provided, propose a short, descriptive name." & vbLf
    s = s & "" & vbLf
    s = s & "3. **Draft an informative `description`**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Write a concise paragraph stating:" & vbLf
    s = s & "     - What the skill does." & vbLf
    s = s & "     - When agents should select it." & vbLf
    s = s & "   - Include domain keywords that match likely user phrasing." & vbLf
    s = s & "" & vbLf
    s = s & "4. **Decide on optional fields**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Add `license` only if the user mentions licensing or the context suggests one." & vbLf
    s = s & "   - Add `compatibility` if the skill depends on specific environments or tools." & vbLf
    s = s & "   - Add `metadata` if the user wants tracking, ownership, categorization, or versions." & vbLf
    s = s & "   - Add `allowed-tools` if there is a known tool policy that must be encoded." & vbLf
    s = s & "" & vbLf
    s = s & "5. **Outline the body**" & vbLf
    s = s & "" & vbLf
    s = s & "   Use a simple, predictable structure such as:" & vbLf
    s = s & "" & vbLf
    s = s & "   ```markdown" & vbLf
    s = s & "   # skill-name" & vbLf
    s = s & "" & vbLf
    s = s & "   ## Purpose" & vbLf
    s = s & "" & vbLf
    s = s & "   ..." & vbLf
    s = s & "" & vbLf
    s = s & "   ## When to use this skill" & vbLf
    s = s & "" & vbLf
    s = s & "   ..." & vbLf
    s = s & "" & vbLf
    s = s & "   ## Required inputs" & vbLf
    s = s & "" & vbLf
    s = s & "   ..." & vbLf
    s = s & "" & vbLf
    s = s & "   ## Workflow" & vbLf
    s = s & "" & vbLf
    s = s & "   1. ..." & vbLf
    s = s & "   2. ..." & vbLf
    s = s & "   3. ..." & vbLf
    s = s & "" & vbLf
    s = s & "   ## Examples" & vbLf
    s = s & "" & vbLf
    s = s & "   ..." & vbLf
    s = s & "" & vbLf
    s = s & "   ## Edge cases and limitations" & vbLf
    s = s & "" & vbLf
    s = s & "   ..." & vbLf
    s = s & "   ```" & vbLf
    s = s & "" & vbLf
    s = s & "6. **Fill in the body with domain-specific guidance**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Translate the user's requirements into clear, stepwise instructions." & vbLf
    s = s & "   - Include at least one example that illustrates typical use." & vbLf
    s = s & "   - If the skill may interact with other skills, note that explicitly." & vbLf
    s = s & "" & vbLf
    s = s & "7. **Suggest supporting files (optional)**" & vbLf
    s = s & "" & vbLf
    s = s & "   If the description implies complex logic or lengthy reference material:" & vbLf
    s = s & "" & vbLf
    s = s & "   - Suggest offloading detailed guidance to:" & vbLf
    s = s & "     - `references/REFERENCE.md` (technical details)." & vbLf
    s = s & "     - `references/FORMS.md` (templates or structured formats)." & vbLf
    s = s & "   - Reference these files from `SKILL.md` using relative paths." & vbLf
    s = s & "" & vbLf
    s = s & "8. **Run an internal validation check**" & vbLf
    s = s & "" & vbLf
    s = s & "   Before returning `SKILL.md`, mentally validate:" & vbLf
    s = s & "" & vbLf
    s = s & "   - `name` is compliant with all rules." & vbLf
    s = s & "   - `description` is non-empty, informative, and under 1024 characters." & vbLf
    s = s & "   - Optional fields, if present, are concise and relevant." & vbLf
    s = s & "   - Body is coherent, formatted as Markdown, and under about 500 lines." & vbLf
    s = s & "   - File references, if any, are one level deep and use relative paths." & vbLf
    s = s & "" & vbLf
    s = s & "9. **Return the final `SKILL.md`**" & vbLf
    s = s & "" & vbLf
    s = s & "   - Present the entire file as one complete Markdown document." & vbLf
    s = s & "   - Do not prepend or append explanatory text outside the file, unless the user explicitly asks for commentary." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 8. Minimal template you can adapt" & vbLf
    s = s & "" & vbLf
    s = s & "When the user wants a **minimal** skill definition, you can adapt this pattern:" & vbLf
    s = s & "" & vbLf
    s = s & "```markdown" & vbLf
    s = s & "---" & vbLf
    s = s & "name: example-skill" & vbLf
    s = s & "description: Concise explanation of what the skill does and when to use it, including relevant domain keywords." & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "# example-skill" & vbLf
    s = s & "" & vbLf
    s = s & "## Purpose" & vbLf
    s = s & "" & vbLf
    s = s & "Briefly describe the main goal of this skill and the types of tasks it supports." & vbLf
    s = s & "" & vbLf
    s = s & "## When to use this skill" & vbLf
    s = s & "" & vbLf
    s = s & "- Bullet 1 describing a typical triggering scenario." & vbLf
    s = s & "- Bullet 2 describing another common scenario." & vbLf
    s = s & "" & vbLf
    s = s & "## Workflow" & vbLf
    s = s & "" & vbLf
    s = s & "1. Confirm the user's goal and key constraints." & vbLf
    s = s & "2. Collect any required inputs (files, parameters, context)." & vbLf
    s = s & "3. Perform the core operation, step by step." & vbLf
    s = s & "4. Summarize results and suggest next steps." & vbLf
    s = s & "" & vbLf
    s = s & "## Examples" & vbLf
    s = s & "" & vbLf
    s = s & "### Example 1" & vbLf
    s = s & "" & vbLf
    s = s & "- **User goal:** ..." & vbLf
    s = s & "- **Agent behavior using this skill:** ..." & vbLf
    s = s & "```" & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 9. Extended template with optional fields" & vbLf
    s = s & "" & vbLf
    s = s & "When a richer definition is requested, you can extend the frontmatter and body:" & vbLf
    s = s & "" & vbLf
    s = s & "```markdown" & vbLf
    s = s & "---" & vbLf
    s = s & "name: example-skill" & vbLf
    s = s & "description: Clear description of what the skill does (its capabilities and operations) and when agents should select it, using domain-relevant keywords." & vbLf
    s = s & "license: Apache-2.0" & vbLf
    s = s & "compatibility: Designed for environments with access to standard command-line tools and local file operations." & vbLf
    s = s & "metadata:" & vbLf
    s = s & "  author: example-org" & vbLf
    s = s & "  version: ""1.0""" & vbLf
    s = s & "  category: ""example-category""" & vbLf
    s = s & "allowed-tools: Bash(git:*) Bash(jq:*) Read" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "# example-skill" & vbLf
    s = s & "" & vbLf
    s = s & "## Purpose" & vbLf
    s = s & "" & vbLf
    s = s & "Explain the main objectives of this skill and the value it provides." & vbLf
    s = s & "" & vbLf
    s = s & "## When to use this skill" & vbLf
    s = s & "" & vbLf
    s = s & "- Scenario 1 where the skill is appropriate." & vbLf
    s = s & "- Scenario 2 where it clearly applies." & vbLf
    s = s & "- Note any scenarios where the skill should *not* be used." & vbLf
    s = s & "" & vbLf
    s = s & "## Required inputs" & vbLf
    s = s & "" & vbLf
    s = s & "List the information, files, or parameters the agent should obtain before proceeding." & vbLf
    s = s & "" & vbLf
    s = s & "## Workflow" & vbLf
    s = s & "" & vbLf
    s = s & "1. Analyze the user's request and match it to the supported operations." & vbLf
    s = s & "2. Verify that all required inputs are present." & vbLf
    s = s & "3. Execute or simulate the relevant steps." & vbLf
    s = s & "4. Handle edge cases or errors gracefully (for example, missing data)." & vbLf
    s = s & "5. Present results with clear explanations and suggested next steps." & vbLf
    s = s & "" & vbLf
    s = s & "## Examples" & vbLf
    s = s & "" & vbLf
    s = s & "Provide one or more examples showing typical usage and outputs." & vbLf
    s = s & "" & vbLf
    s = s & "## Edge cases and limitations" & vbLf
    s = s & "" & vbLf
    s = s & "Describe ambiguous cases, unsupported operations, or constraints the agent must respect." & vbLf
    s = s & "" & vbLf
    s = s & "## Additional references" & vbLf
    s = s & "" & vbLf
    s = s & "If detailed technical information exists, mention where it lives, for example:" & vbLf
    s = s & "" & vbLf
    s = s & "- See `references/REFERENCE.md` for a detailed technical guide." & vbLf
    s = s & "```" & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "## 10. Validation checklist (for every SKILL.md you produce)" & vbLf
    s = s & "" & vbLf
    s = s & "Before finalizing any `SKILL.md`, confirm that:" & vbLf
    s = s & "" & vbLf
    s = s & "1. **Frontmatter structure**" & vbLf
    s = s & "   - Starts with `---` on its own line." & vbLf
    s = s & "   - Contains at least `name` and `description`." & vbLf
    s = s & "   - Ends with `---` on its own line before the body." & vbLf
    s = s & "" & vbLf
    s = s & "2. **`name` field**" & vbLf
    s = s & "   - 1-64 characters." & vbLf
    s = s & "   - Only lowercase letters, digits, and hyphens." & vbLf
    s = s & "   - No leading, trailing, or consecutive hyphens." & vbLf
    s = s & "" & vbLf
    s = s & "3. **`description` field**" & vbLf
    s = s & "   - Non-empty and under 1024 characters." & vbLf
    s = s & "   - Explains what the skill does and when to use it." & vbLf
    s = s & "   - Contains relevant domain keywords." & vbLf
    s = s & "" & vbLf
    s = s & "4. **Optional fields**" & vbLf
    s = s & "   - Present only when justified." & vbLf
    s = s & "   - Values are concise and clearly written." & vbLf
    s = s & "   - `metadata` keys and values are strings." & vbLf
    s = s & "" & vbLf
    s = s & "5. **Body content**" & vbLf
    s = s & "   - Begins after the closing `---` of the frontmatter." & vbLf
    s = s & "   - Uses valid Markdown headings and lists." & vbLf
    s = s & "   - Provides actionable guidance, examples, and edge case notes." & vbLf
    s = s & "   - Keeps very detailed reference material out of the main body when it would be too long." & vbLf
    s = s & "" & vbLf
    s = s & "6. **File references**" & vbLf
    s = s & "   - Use relative paths from the skill root." & vbLf
    s = s & "   - Are at most one level deep (e.g., `references/REFERENCE.md`)." & vbLf
    s = s & "" & vbLf
    s = s & "If any checklist item fails, revise the draft before returning it." & vbLf
    s = s & "" & vbLf
    s = s & "---" & vbLf
    SeedSkillMdAuthoring = s
End Function


