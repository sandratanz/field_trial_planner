# ============================================================
# Field Trial Planner (GRDC-West)
# ------------------------------------------------------------
# A Shiny data-collection tool that walks a grower/researcher
# through the AAGI Analytics Collaboration Plan (ACP): project
# overview, trials, factors & levels, responses, research
# questions, and implementation sites. The completed plan can
# be downloaded as a structured Excel workbook for AAGI to use
# when designing the trial layout.
#
# Required packages:
#   install.packages(c("shiny", "DT", "openxlsx"))
#
# Run with:
#   shiny::runApp("app.R")
# ============================================================

library(shiny)
library(DT)
library(openxlsx)

# ------------------------------------------------------------
# Reference choice lists (from the AAGI ACP "Lists" sheet)
# ------------------------------------------------------------

layout_choices <- c(
  "RCBD", "Split-plot", "Split-split plot", "Strip plot",
  "Alpha design", "Augmented design", "Row-column design",
  "To be designed by AAGI"
)

data_type_choices <- c(
  "Continuous", "Count", "Proportion", "Ordinal", "Binary", "Time-to-event"
)

effect_type_choices <- c("Main effect", "Two-way interaction", "Three-way interaction")

want_to_know_choices <- c(
  "Difference between levels", "Optimum or dose-response",
  "Threshold", "Stability across environments"
)

design_support_choices   <- c("Not required", "Review existing design", "AAGI to design", "Unsure")
analysis_support_choices <- c("Not required", "Required", "Unsure")

# ------------------------------------------------------------
# Empty data frame constructors (define column types up front
# so rbind() and Excel export behave predictably even when a
# table has zero rows)
# ------------------------------------------------------------

empty_trials <- function() data.frame(
  trial_id = character(), trial_name = character(), trial_aim = character(),
  design_support = character(), analysis_support = character(),
  stringsAsFactors = FALSE
)

empty_factors <- function() data.frame(
  trial_id = character(), factor_name = character(), n_levels = integer(),
  stringsAsFactors = FALSE
)

empty_levels <- function() data.frame(
  trial_id = character(), factor_name = character(), level_num = integer(),
  level_desc = character(), stringsAsFactors = FALSE
)

empty_responses <- function() data.frame(
  trial_id = character(), response_var = character(), units = character(),
  data_type = character(), measurement_method = character(),
  sampling = character(), repeated_measures = character(),
  timing = character(), stringsAsFactors = FALSE
)

empty_questions <- function() data.frame(
  trial_id = character(), response_var = character(), effect_type = character(),
  factors_involved = character(), want_to_know = character(),
  question_text = character(), stringsAsFactors = FALSE
)

empty_implementation <- function() data.frame(
  trial_id = character(), location = character(), year = character(),
  contact = character(), n_reps = character(), layout = character(),
  notes = character(), stringsAsFactors = FALSE
)

empty_summary <- function() data.frame(
  trial_id = character(), trial_name = character(), implementations = integer(),
  designs = integer(), analyses = integer(), multi_factor = logical(),
  interaction = logical(), multi_site = logical(), nonstandard_response = logical(),
  split_strip_layout = logical(), complexity = character(), design_support = character(),
  analysis_support = character(), attention = character(), complexity_drivers = character(),
  stringsAsFactors = FALSE
)

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

# Null-coalesce, used when restoring a saved progress file that may
# predate a field, or where a field was left blank.
`%||%` <- function(a, b) if (is.null(a)) b else a

# How many factors a given effect type needs selected
n_factors_needed <- function(effect_type) {
  switch(effect_type,
    "Main effect" = 1,
    "Two-way interaction" = 2,
    "Three-way interaction" = 3,
    1
  )
}

# Turns (effect type, factors, "what you want to know", response) into a
# plain-English research question, mirroring the worked examples in the
# AAGI ACP template.
generate_question <- function(effect_type, factors_involved, want_to_know, response) {
  factors_txt <- gsub(",\\s*", " and ", factors_involved)
  multi <- grepl("interaction", effect_type, ignore.case = TRUE)

  if (want_to_know == "Difference between levels") {
    if (multi) sprintf("Does the interaction between %s affect %s?", factors_txt, response)
    else sprintf("Does %s affect %s?", factors_txt, response)
  } else if (want_to_know == "Optimum or dose-response") {
    if (multi) sprintf("What combination of %s maximises %s?", factors_txt, response)
    else sprintf("What level of %s maximises %s?", factors_txt, response)
  } else if (want_to_know == "Threshold") {
    if (multi) sprintf("At what combination of %s does %s cross the target threshold?", factors_txt, response)
    else sprintf("At what level of %s does %s cross the target threshold?", factors_txt, response)
  } else if (want_to_know == "Stability across environments") {
    if (multi) sprintf("Is the interaction between %s on %s stable across environments?", factors_txt, response)
    else sprintf("Is the effect of %s on %s stable across environments?", factors_txt, response)
  } else {
    sprintf("How does %s affect %s?", factors_txt, response)
  }
}

# ------------------------------------------------------------
# UI section builders
# ------------------------------------------------------------

project_overview_ui <- function() {
  tagList(
    p(class = "section-help",
      "Complete this section once for the whole project. It gives AAGI the context for every trial described in later sections."),
    fluidRow(
      column(6,
        textInput("proj_title", "Project title",
                  placeholder = "e.g. Optimising nitrogen and sowing strategies for wheat productivity in the WA central grainbelt"),
        textInput("proj_id", "Project number / ID", placeholder = "e.g. WGRG-2026-014"),
        textInput("proj_pi", "Principal investigator / research group",
                  placeholder = "Name(s) and contact details of the lead researcher(s) or research group")
      ),
      column(6,
        textInput("proj_form_by", "Form completed by", placeholder = "If different from the lead researcher"),
        textInput("proj_email", "Contact email", placeholder = "Best email address for questions about this form"),
        textAreaInput("proj_objective", "Overall project objective", rows = 4,
                      placeholder = "The broad goal or intended outcome of the project")
      )
    )
  )
}

trials_ui <- function() {
  tagList(
    p(class = "section-help",
      "Each row is one trial design. If the same design runs at multiple sites or years, list it once here \u2014 you'll record each implementation in Section 2.6. Create a separate trial when the research question, treatments, or design differ."),
    wellPanel(
      fluidRow(
        column(4, textInput("trial_name_in", "Trial name", placeholder = "e.g. Nitrogen rate trial")),
        column(8, textAreaInput("trial_aim_in", "Trial aim", rows = 2,
                                 placeholder = "What this trial is investigating \u2014 the effect or relationship being examined and the key factors central to it."))
      ),
      fluidRow(
        column(4, selectInput("trial_design_support_in", "Design support requested", choices = design_support_choices)),
        column(4, selectInput("trial_analysis_support_in", "Analysis support requested", choices = analysis_support_choices)),
        column(4, br(), actionButton("add_trial", "Add trial", class = "btn-primary", icon = icon("plus")))
      )
    ),
    div(class = "table-block",
        DTOutput("trials_table"), br(),
        actionButton("remove_trial", "Remove selected trial(s)", class = "btn-danger")
    )
  )
}

factors_ui <- function() {
  tagList(
    p(class = "section-help",
      "List the treatment factors for each trial, then describe every level. Keep the number of treatment combinations manageable \u2014 more factors and levels multiply the number of plots required."),
    wellPanel(
      fluidRow(
        column(4, uiOutput("factor_trial_select")),
        column(4, textInput("factor_name_in", "Factor name", placeholder = "e.g. Nitrogen rate")),
        column(2, numericInput("factor_nlevels_in", "Number of levels", value = 2, min = 2, max = 12, step = 1)),
        column(2, br(), actionButton("add_factor", "Add factor", class = "btn-primary", icon = icon("plus")))
      )
    ),
    div(class = "table-block",
        h4("Table 1: Factors"),
        DTOutput("factors_table"), br(),
        actionButton("remove_factor", "Remove selected factor(s)", class = "btn-danger")
    ),
    hr(),
    h4("Table 2: Level descriptions"),
    p(class = "section-help",
      "Describe each level fully enough that someone else could implement it \u2014 product, rate and timing where relevant. Fill this in once your factor list (above) is finalised, then click Save."),
    uiOutput("level_inputs"),
    actionButton("save_levels", "Save level descriptions", class = "btn-primary"),
    br(), br(),
    DTOutput("levels_table")
  )
}

responses_ui <- function() {
  tagList(
    p(class = "section-help",
      "List every response variable you will measure, one per row. Each distinct response is a separate analysis."),
    wellPanel(
      fluidRow(
        column(4, uiOutput("response_trial_select")),
        column(4, textInput("response_var_in", "Response variable", placeholder = "e.g. Grain yield")),
        column(4, textInput("response_units_in", "Units", placeholder = "e.g. t/ha"))
      ),
      fluidRow(
        column(4, selectInput("response_datatype_in", "Data type", choices = data_type_choices)),
        column(4, selectInput("response_repeated_in", "Repeated measures?", choices = c("No", "Yes"))),
        column(4, textInput("response_timing_in", "Measurement timing", placeholder = "e.g. Once at harvest"))
      ),
      fluidRow(
        column(6, textAreaInput("response_method_in", "Measurement method", rows = 2,
                                 placeholder = "How the outcome is recorded, including instrument or scoring system")),
        column(6, textAreaInput("response_sampling_in", "Sampling within experimental unit", rows = 2,
                                 placeholder = "Describe the actual procedure, not just that it is random"))
      ),
      actionButton("add_response", "Add response", class = "btn-primary", icon = icon("plus"))
    ),
    div(class = "table-block",
        DTOutput("responses_table"), br(),
        actionButton("remove_response", "Remove selected response(s)", class = "btn-danger")
    )
  )
}

questions_ui <- function() {
  tagList(
    p(class = "section-help",
      "Build each research question from a response, the effect of interest, and the kind of answer you need. Only factors and responses already listed for the chosen trial are available."),
    wellPanel(
      fluidRow(
        column(3, uiOutput("question_trial_select")),
        column(3, uiOutput("question_response_select")),
        column(3, selectInput("question_effect_in", "Effect type", choices = effect_type_choices)),
        column(3, selectInput("question_want_in", "What you want to know", choices = want_to_know_choices))
      ),
      fluidRow(
        column(9, uiOutput("question_factors_select")),
        column(3, br(), actionButton("add_question", "Add question", class = "btn-primary", icon = icon("plus")))
      ),
      uiOutput("question_preview")
    ),
    div(class = "table-block",
        DTOutput("questions_table"), br(),
        actionButton("remove_question", "Remove selected question(s)", class = "btn-danger")
    )
  )
}

implementation_ui <- function() {
  tagList(
    p(class = "section-help",
      "Record every unique implementation of each trial \u2014 one row per site x year combination."),
    wellPanel(
      fluidRow(
        column(3, uiOutput("impl_trial_select")),
        column(3, textInput("impl_location_in", "Location", placeholder = "e.g. Merredin")),
        column(2, textInput("impl_year_in", "Year", placeholder = "e.g. 2027")),
        column(4, textInput("impl_contact_in", "Implementation contact", placeholder = "Local contact responsible for implementation"))
      ),
      fluidRow(
        column(2, numericInput("impl_reps_in", "Replicates", value = 4, min = 1, step = 1)),
        column(4, selectInput("impl_layout_in", "Experimental layout", choices = layout_choices)),
        column(6, textAreaInput("impl_notes_in", "Notes or deviations", rows = 2,
                                 placeholder = "Anything that differed from the plan, or conditions affecting the data"))
      ),
      actionButton("add_impl", "Add implementation", class = "btn-primary", icon = icon("plus"))
    ),
    div(class = "table-block",
        DTOutput("implementation_table"), br(),
        actionButton("remove_impl", "Remove selected implementation(s)", class = "btn-danger")
    )
  )
}

summary_ui <- function() {
  tagList(
    p(class = "section-help",
      "This section fills in automatically from Sections 2.1\u20132.6, and summarises the scale and complexity of what you have described \u2014 useful for both you and AAGI to see where design or analysis support is likely to be needed."),
    DTOutput("summary_table")
  )
}

roles_ui <- function() {
  tagList(
    h4("Researcher responsibilities"),
    p("Data provision, metadata and contextual details for all trials, including detailed trial protocols and visual representations of all trial layouts. Documentation must be kept up to date and record any changes made during implementation, so that what AAGI receives correctly reflects the data being analysed."),
    h5("Information required for experimental layout design"),
    tags$ul(
      tags$li("Research question(s), objectives and hypotheses the trial is intended to address."),
      tags$li("Treatments or factors to be evaluated, including components and application requirements."),
      tags$li("Planned number of sites and available trial area at each site, including spatial or operational constraints."),
      tags$li("Relevant site characteristics, including known gradients or sources of spatial variation."),
      tags$li("Practical constraints affecting implementation \u2014 machinery access, buffer requirements, plot dimensions.")
    ),
    h5("Information required for data analysis"),
    tags$ul(
      tags$li("Specific research question(s) and objectives the analysis is intended to address."),
      tags$li("Experimental design and trial layout, including sites, replication structure, and visualisations."),
      tags$li("Trial protocol, including treatment structure, sampling approach and measurement methods."),
      tags$li("Trial dataset, including a data dictionary describing all variables."),
      tags$li("Relevant site, implementation or environmental notes.")
    ),
    h4("Analyst (AAGI) responsibilities"),
    tags$ul(
      tags$li("Design appropriate, feasible, analytics-informed trial layouts for the specified research questions."),
      tags$li("Plan and perform data analyses based on the information provided."),
      tags$li("Deliver technical outputs (design layouts and data analysis)."),
      tags$li("Answer questions and provide analytics insight and guidance for implementation and interpretation.")
    ),
    h4("Outputs provided by AAGI"),
    tags$ul(
      tags$li("An experimental design layout showing the spatial arrangement of plots and treatment allocations."),
      tags$li("A design listing containing plot-level design information (plot IDs, row/column positions, treatment allocations)."),
      tags$li("A technical report summarising the statistical analysis, results, tables, figures and model outputs.")
    ),
    hr(),
    h4("Sign-off"),
    fluidRow(
      column(3, textInput("signoff_name", "Name")),
      column(3, textInput("signoff_role", "Role")),
      column(3, textInput("signoff_date", "Date", value = as.character(Sys.Date())))
    )
  )
}

download_ui <- function() {
  tagList(
    p("Download the completed Analytics Collaboration Plan as an Excel workbook, structured to match AAGI's ACP format, ready to send to your AAGI contact."),
    downloadButton("download_acp", "Download ACP (.xlsx)", class = "btn-success btn-lg")
  )
}

save_load_ui <- function() {
  tagList(
    p("Save your progress to a file so you can close the app and pick up again later, or hand an in-progress plan to a colleague to continue."),
    fluidRow(
      column(6,
        wellPanel(
          h4("Save progress"),
          p(class = "section-help", "Downloads a single file with everything entered so far across all tabs."),
          downloadButton("save_progress", "Save progress (.rds)", class = "btn-primary")
        )
      ),
      column(6,
        wellPanel(
          h4("Load progress"),
          p(class = "section-help", "Choose a previously saved .rds file to restore your plan. This replaces anything currently in the form."),
          fileInput("load_progress_file", NULL, accept = ".rds",
                    buttonLabel = "Browse...", placeholder = "No file selected")
        )
      )
    )
  )
}

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------

ui <- fluidPage(
  title = "Field Trial Planner (GRDC-West)",
  tags$head(tags$style(HTML("
    .app-header { padding: 18px 0 8px 0; }
    .app-header h2 { margin-bottom: 2px; }
    .app-header p { color: #555; margin-top: 0; }
    .section-help { color: #666; font-size: 13px; margin-bottom: 14px; }
    .table-block { margin-top: 18px; }
    .well { background-color: #fafafa; }
  "))),
  div(class = "app-header",
      h2("Field Trial Planner"),
      p("GRDC-West \u2014 Analytics Collaboration Plan (ACP) intake tool. Complete each section, then download the completed plan for AAGI.")
  ),
  tabsetPanel(
    id = "main_tabs",
    tabPanel("1. Project Overview", br(), project_overview_ui()),
    tabPanel("2.1 Trials", br(), trials_ui()),
    tabPanel("2.2 Factors & Levels", br(), factors_ui()),
    tabPanel("2.4 Responses", br(), responses_ui()),
    tabPanel("2.5 Questions", br(), questions_ui()),
    tabPanel("2.6 Implementation", br(), implementation_ui()),
    tabPanel("2.7 Summary", br(), summary_ui()),
    tabPanel("Roles & Sign-off", br(), roles_ui()),
    tabPanel("Save, Load & Download", br(), save_load_ui(), hr(), download_ui())
  )
)

# ------------------------------------------------------------
# Server
# ------------------------------------------------------------

server <- function(input, output, session) {

  rv <- reactiveValues(
    trials = empty_trials(),
    factors = empty_factors(),
    levels = empty_levels(),
    responses = empty_responses(),
    questions = empty_questions(),
    implementation = empty_implementation()
  )

  trial_choices <- reactive({
    if (nrow(rv$trials) == 0) return(character())
    setNames(rv$trials$trial_id, paste0(rv$trials$trial_id, " - ", rv$trials$trial_name))
  })

  # ---------------- 2.1 Trials ----------------
  observeEvent(input$add_trial, {
    req(input$trial_name_in)
    new_id <- paste0("T", nrow(rv$trials) + 1)
    rv$trials <- rbind(rv$trials, data.frame(
      trial_id = new_id,
      trial_name = input$trial_name_in,
      trial_aim = input$trial_aim_in,
      design_support = input$trial_design_support_in,
      analysis_support = input$trial_analysis_support_in,
      stringsAsFactors = FALSE
    ))
    updateTextInput(session, "trial_name_in", value = "")
    updateTextAreaInput(session, "trial_aim_in", value = "")
  })

  output$trials_table <- renderDT({
    datatable(rv$trials, selection = "multiple", rownames = FALSE,
              colnames = c("Trial ID", "Trial Name", "Trial Aim", "Design Support", "Analysis Support"),
              options = list(dom = "tp", pageLength = 10))
  })

  observeEvent(input$remove_trial, {
    sel <- input$trials_table_rows_selected
    req(length(sel) > 0)
    removed_ids <- rv$trials$trial_id[sel]
    rv$trials <- rv$trials[-sel, , drop = FALSE]
    rv$factors <- rv$factors[!rv$factors$trial_id %in% removed_ids, , drop = FALSE]
    rv$levels <- rv$levels[!rv$levels$trial_id %in% removed_ids, , drop = FALSE]
    rv$responses <- rv$responses[!rv$responses$trial_id %in% removed_ids, , drop = FALSE]
    rv$questions <- rv$questions[!rv$questions$trial_id %in% removed_ids, , drop = FALSE]
    rv$implementation <- rv$implementation[!rv$implementation$trial_id %in% removed_ids, , drop = FALSE]
  })

  # ---------------- 2.2 Factors & Levels ----------------
  output$factor_trial_select <- renderUI({
    selectInput("factor_trial_in", "Trial", choices = trial_choices())
  })

  observeEvent(input$add_factor, {
    req(input$factor_trial_in, input$factor_name_in)
    rv$factors <- rbind(rv$factors, data.frame(
      trial_id = input$factor_trial_in,
      factor_name = input$factor_name_in,
      n_levels = input$factor_nlevels_in,
      stringsAsFactors = FALSE
    ))
    updateTextInput(session, "factor_name_in", value = "")
  })

  output$factors_table <- renderDT({
    df <- rv$factors
    if (nrow(df) > 0) {
      df$trial_label <- vapply(df$trial_id, function(id) {
        nm <- rv$trials$trial_name[rv$trials$trial_id == id]
        if (length(nm) == 0) id else paste0(id, " - ", nm)
      }, character(1))
      df <- df[, c("trial_label", "factor_name", "n_levels")]
    }
    datatable(df, selection = "multiple", rownames = FALSE,
              colnames = c("Trial", "Factor", "Number of Levels"),
              options = list(dom = "tp", pageLength = 10))
  })

  observeEvent(input$remove_factor, {
    sel <- input$factors_table_rows_selected
    req(length(sel) > 0)
    removed <- rv$factors[sel, , drop = FALSE]
    rv$factors <- rv$factors[-sel, , drop = FALSE]
    for (i in seq_len(nrow(removed))) {
      rv$levels <- rv$levels[!(rv$levels$trial_id == removed$trial_id[i] &
                                  rv$levels$factor_name == removed$factor_name[i]), , drop = FALSE]
    }
  })

  output$level_inputs <- renderUI({
    if (nrow(rv$factors) == 0) return(p(em("Add at least one factor above to describe its levels.")))
    panels <- lapply(seq_len(nrow(rv$factors)), function(i) {
      fr <- rv$factors[i, ]
      trial_nm <- rv$trials$trial_name[rv$trials$trial_id == fr$trial_id]
      label <- paste0(fr$trial_id, " - ", if (length(trial_nm)) trial_nm else "", " : ", fr$factor_name)
      level_inputs <- lapply(seq_len(fr$n_levels), function(j) {
        existing <- rv$levels$level_desc[rv$levels$trial_id == fr$trial_id &
                                            rv$levels$factor_name == fr$factor_name &
                                            rv$levels$level_num == j]
        val <- if (length(existing)) existing[1] else ""
        textInput(paste0("lvl_", i, "_", j), paste0("Level ", j), value = val, width = "100%")
      })
      wellPanel(h5(label), do.call(tagList, level_inputs))
    })
    do.call(tagList, panels)
  })

  observeEvent(input$save_levels, {
    req(nrow(rv$factors) > 0)
    new_levels <- empty_levels()
    for (i in seq_len(nrow(rv$factors))) {
      fr <- rv$factors[i, ]
      for (j in seq_len(fr$n_levels)) {
        val <- input[[paste0("lvl_", i, "_", j)]]
        new_levels <- rbind(new_levels, data.frame(
          trial_id = fr$trial_id, factor_name = fr$factor_name,
          level_num = j, level_desc = ifelse(is.null(val), "", val),
          stringsAsFactors = FALSE
        ))
      }
    }
    rv$levels <- new_levels
    showNotification("Level descriptions saved.", type = "message")
  })

  output$levels_table <- renderDT({
    datatable(rv$levels, rownames = FALSE,
              colnames = c("Trial ID", "Factor", "Level #", "Description"),
              options = list(dom = "tp", pageLength = 10))
  })

  # ---------------- 2.4 Responses ----------------
  output$response_trial_select <- renderUI({
    selectInput("response_trial_in", "Trial", choices = trial_choices())
  })

  observeEvent(input$add_response, {
    req(input$response_trial_in, input$response_var_in)
    rv$responses <- rbind(rv$responses, data.frame(
      trial_id = input$response_trial_in,
      response_var = input$response_var_in,
      units = input$response_units_in,
      data_type = input$response_datatype_in,
      measurement_method = input$response_method_in,
      sampling = input$response_sampling_in,
      repeated_measures = input$response_repeated_in,
      timing = input$response_timing_in,
      stringsAsFactors = FALSE
    ))
    updateTextInput(session, "response_var_in", value = "")
    updateTextInput(session, "response_units_in", value = "")
    updateTextAreaInput(session, "response_method_in", value = "")
    updateTextAreaInput(session, "response_sampling_in", value = "")
    updateTextInput(session, "response_timing_in", value = "")
  })

  output$responses_table <- renderDT({
    datatable(rv$responses, selection = "multiple", rownames = FALSE,
              colnames = c("Trial ID", "Response Variable", "Units", "Data Type",
                           "Measurement Method", "Sampling", "Repeated?", "Timing"),
              options = list(dom = "tp", pageLength = 10, scrollX = TRUE))
  })

  observeEvent(input$remove_response, {
    sel <- input$responses_table_rows_selected
    req(length(sel) > 0)
    removed <- rv$responses[sel, , drop = FALSE]
    rv$responses <- rv$responses[-sel, , drop = FALSE]
    for (i in seq_len(nrow(removed))) {
      rv$questions <- rv$questions[!(rv$questions$trial_id == removed$trial_id[i] &
                                        rv$questions$response_var == removed$response_var[i]), , drop = FALSE]
    }
  })

  # ---------------- 2.5 Questions ----------------
  output$question_trial_select <- renderUI({
    selectInput("question_trial_in", "Trial", choices = trial_choices())
  })

  output$question_response_select <- renderUI({
    req(input$question_trial_in)
    choices <- rv$responses$response_var[rv$responses$trial_id == input$question_trial_in]
    selectInput("question_response_in", "Response variable", choices = choices)
  })

  output$question_factors_select <- renderUI({
    req(input$question_trial_in)
    choices <- rv$factors$factor_name[rv$factors$trial_id == input$question_trial_in]
    n_needed <- n_factors_needed(input$question_effect_in)
    selectizeInput("question_factors_in",
                    paste0("Factor(s) involved (choose ", n_needed, ")"),
                    choices = choices, multiple = TRUE,
                    options = list(maxItems = n_needed))
  })

  question_preview_text <- reactive({
    req(input$question_response_in, input$question_factors_in,
        input$question_want_in, input$question_effect_in)
    n_needed <- n_factors_needed(input$question_effect_in)
    if (length(input$question_factors_in) != n_needed) return(NULL)
    generate_question(input$question_effect_in, paste(input$question_factors_in, collapse = ", "),
                       input$question_want_in, input$question_response_in)
  })

  output$question_preview <- renderUI({
    txt <- question_preview_text()
    if (is.null(txt)) {
      p(em("Choose a trial, response, effect type and the right number of factors to preview the question."))
    } else {
      div(style = "margin-top:10px;padding:10px;background:#eef6ff;border-radius:4px;",
          strong("Research question: "), txt)
    }
  })

  observeEvent(input$add_question, {
    txt <- question_preview_text()
    req(txt, input$question_trial_in)
    rv$questions <- rbind(rv$questions, data.frame(
      trial_id = input$question_trial_in,
      response_var = input$question_response_in,
      effect_type = input$question_effect_in,
      factors_involved = paste(input$question_factors_in, collapse = ", "),
      want_to_know = input$question_want_in,
      question_text = txt,
      stringsAsFactors = FALSE
    ))
  })

  output$questions_table <- renderDT({
    datatable(rv$questions, selection = "multiple", rownames = FALSE,
              colnames = c("Trial ID", "Response", "Effect Type", "Factor(s)",
                           "What You Want to Know", "Research Question"),
              options = list(dom = "tp", pageLength = 10, scrollX = TRUE))
  })

  observeEvent(input$remove_question, {
    sel <- input$questions_table_rows_selected
    req(length(sel) > 0)
    rv$questions <- rv$questions[-sel, , drop = FALSE]
  })

  # ---------------- 2.6 Implementation ----------------
  output$impl_trial_select <- renderUI({
    selectInput("impl_trial_in", "Trial", choices = trial_choices())
  })

  observeEvent(input$add_impl, {
    req(input$impl_trial_in, input$impl_location_in, input$impl_year_in)
    rv$implementation <- rbind(rv$implementation, data.frame(
      trial_id = input$impl_trial_in,
      location = input$impl_location_in,
      year = input$impl_year_in,
      contact = input$impl_contact_in,
      n_reps = as.character(input$impl_reps_in),
      layout = input$impl_layout_in,
      notes = input$impl_notes_in,
      stringsAsFactors = FALSE
    ))
    updateTextInput(session, "impl_location_in", value = "")
    updateTextInput(session, "impl_year_in", value = "")
    updateTextInput(session, "impl_contact_in", value = "")
    updateTextAreaInput(session, "impl_notes_in", value = "")
  })

  output$implementation_table <- renderDT({
    datatable(rv$implementation, selection = "multiple", rownames = FALSE,
              colnames = c("Trial ID", "Location", "Year", "Contact", "Replicates", "Layout", "Notes"),
              options = list(dom = "tp", pageLength = 10, scrollX = TRUE))
  })

  observeEvent(input$remove_impl, {
    sel <- input$implementation_table_rows_selected
    req(length(sel) > 0)
    rv$implementation <- rv$implementation[-sel, , drop = FALSE]
  })

  # ---------------- 2.7 Summary ----------------
  summary_df <- reactive({
    if (nrow(rv$trials) == 0) return(empty_summary())
    do.call(rbind, lapply(seq_len(nrow(rv$trials)), function(i) {
      tid <- rv$trials$trial_id[i]
      impl <- rv$implementation[rv$implementation$trial_id == tid, , drop = FALSE]
      facs <- rv$factors[rv$factors$trial_id == tid, , drop = FALSE]
      resp <- rv$responses[rv$responses$trial_id == tid, , drop = FALSE]
      ques <- rv$questions[rv$questions$trial_id == tid, , drop = FALSE]

      n_impl <- nrow(impl)
      n_designs <- length(unique(impl$layout))
      n_analyses <- nrow(ques)
      multi_factor <- nrow(facs) > 1
      interaction <- any(grepl("interaction", ques$effect_type, ignore.case = TRUE))
      multi_site <- length(unique(impl$location)) > 1
      nonstd_resp <- any(resp$data_type != "Continuous")
      split_strip <- any(grepl("split|strip", impl$layout, ignore.case = TRUE))

      flags <- c(multi_factor, interaction, multi_site, nonstd_resp, split_strip)
      score <- sum(flags)
      complexity <- if (score >= 3) "High" else if (score >= 1) "Moderate" else "Standard"
      attention <- if (score >= 3) "Review recommended" else "Standard"

      driver_names <- c("Multi-factor", "Interaction", "Multi-site", "Non-standard response", "Split/strip layout")
      drivers <- paste(driver_names[flags], collapse = "; ")
      if (drivers == "") drivers <- "-"

      data.frame(
        trial_id = tid, trial_name = rv$trials$trial_name[i],
        implementations = n_impl, designs = n_designs, analyses = n_analyses,
        multi_factor = multi_factor, interaction = interaction, multi_site = multi_site,
        nonstandard_response = nonstd_resp, split_strip_layout = split_strip,
        complexity = complexity,
        design_support = rv$trials$design_support[i],
        analysis_support = rv$trials$analysis_support[i],
        attention = attention, complexity_drivers = drivers,
        stringsAsFactors = FALSE
      )
    }))
  })

  output$summary_table <- renderDT({
    datatable(summary_df(), rownames = FALSE, options = list(dom = "tp", pageLength = 10, scrollX = TRUE))
  })

  # ---------------- Save / Load progress ----------------
  get_state <- function() {
    list(
      version = 1,
      project = list(
        title = input$proj_title, id = input$proj_id, pi = input$proj_pi,
        form_by = input$proj_form_by, email = input$proj_email, objective = input$proj_objective
      ),
      trials = rv$trials,
      factors = rv$factors,
      levels = rv$levels,
      responses = rv$responses,
      questions = rv$questions,
      implementation = rv$implementation,
      signoff = list(name = input$signoff_name, role = input$signoff_role, date = input$signoff_date)
    )
  }

  output$save_progress <- downloadHandler(
    filename = function() {
      title <- input$proj_title
      safe <- if (!is.null(title) && nzchar(title)) gsub("[^A-Za-z0-9]+", "_", title) else "Field_Trial_Plan"
      paste0("FTP_progress_", safe, "_", format(Sys.Date(), "%Y%m%d"), ".rds")
    },
    content = function(file) {
      saveRDS(get_state(), file)
    }
  )

  observeEvent(input$load_progress_file, {
    req(input$load_progress_file)
    state <- tryCatch(readRDS(input$load_progress_file$datapath), error = function(e) NULL)
    if (is.null(state)) {
      showNotification("Could not read that file \u2014 is it a Field Trial Planner .rds save file?", type = "error")
      return(invisible())
    }

    rv$trials <- state$trials %||% empty_trials()
    rv$factors <- state$factors %||% empty_factors()
    rv$levels <- state$levels %||% empty_levels()
    rv$responses <- state$responses %||% empty_responses()
    rv$questions <- state$questions %||% empty_questions()
    rv$implementation <- state$implementation %||% empty_implementation()

    proj <- state$project %||% list()
    updateTextInput(session, "proj_title", value = proj$title %||% "")
    updateTextInput(session, "proj_id", value = proj$id %||% "")
    updateTextInput(session, "proj_pi", value = proj$pi %||% "")
    updateTextInput(session, "proj_form_by", value = proj$form_by %||% "")
    updateTextInput(session, "proj_email", value = proj$email %||% "")
    updateTextAreaInput(session, "proj_objective", value = proj$objective %||% "")

    so <- state$signoff %||% list()
    updateTextInput(session, "signoff_name", value = so$name %||% "")
    updateTextInput(session, "signoff_role", value = so$role %||% "")
    updateTextInput(session, "signoff_date", value = so$date %||% as.character(Sys.Date()))

    showNotification("Progress loaded.", type = "message")
  })

  # ---------------- Download ----------------
  output$download_acp <- downloadHandler(
    filename = function() {
      title <- input$proj_title
      safe <- if (!is.null(title) && nzchar(title)) gsub("[^A-Za-z0-9]+", "_", title) else "Field_Trial_Plan"
      paste0("ACP_", safe, "_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      wb <- createWorkbook()
      hs <- createStyle(textDecoration = "bold", fgFill = "#DCE6F1", border = "Bottom")

      write_sheet <- function(name, df, display_names = NULL) {
        addWorksheet(wb, name)
        if (!is.null(display_names) && ncol(df) == length(display_names)) names(df) <- display_names
        writeData(wb, name, df, headerStyle = hs)
        if (ncol(df) > 0) setColWidths(wb, name, cols = 1:ncol(df), widths = "auto")
      }

      po <- data.frame(
        Field = c("Project title", "Project number / ID", "Principal investigator / research group",
                  "Form completed by", "Contact email", "Overall project objective"),
        Response = c(input$proj_title, input$proj_id, input$proj_pi,
                     input$proj_form_by, input$proj_email, input$proj_objective),
        stringsAsFactors = FALSE
      )

      write_sheet("1. Project Overview", po)
      write_sheet("2.1 Trials", rv$trials,
                  c("Trial ID", "Trial Name", "Trial Aim", "Design Support Requested", "Analysis Support Requested"))
      write_sheet("2.2 Factors", rv$factors, c("Trial ID", "Factor", "Number of Levels"))
      write_sheet("2.3 Levels", rv$levels, c("Trial ID", "Factor", "Level #", "Level Description"))
      write_sheet("2.4 Responses", rv$responses,
                  c("Trial ID", "Response Variable", "Units", "Data Type", "Measurement Method",
                    "Sampling Within Experimental Unit", "Repeated Measures?", "Measurement Timing"))
      write_sheet("2.5 Questions", rv$questions,
                  c("Trial ID", "Response Variable", "Effect Type", "Factor(s) Involved",
                    "What You Want to Know", "Research Question"))
      write_sheet("2.6 Implementation", rv$implementation,
                  c("Trial ID", "Location", "Year", "Implementation Contact", "Number of Replicates",
                    "Experimental Layout", "Notes or Deviations"))
      write_sheet("2.7 Summary", summary_df(),
                  c("Trial ID", "Trial Name", "Implementations", "Designs", "Analyses", "Multi-factor",
                    "Interaction", "Multi-site", "Non-standard Response", "Split/Strip Layout",
                    "Complexity", "Design Support Requested", "Analysis Support Requested",
                    "Attention", "Complexity Drivers"))

      so <- data.frame(
        Name = input$signoff_name, Role = input$signoff_role, Date = input$signoff_date,
        stringsAsFactors = FALSE
      )
      write_sheet("3. Sign-off", so)

      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui = ui, server = server)
