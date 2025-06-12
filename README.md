# APA-Style-Tables-in-R
 to generate SPSS-style APA tables directly into a perfectly formatted Word document — with Times New Roman font, APA-compliant layout, and ready-to-use interpretation for research theses and academic reporting."


# 📋 Step 1: Import Data from Clipboard
# ------------------------------------------------
# This line reads tabular data copied from Excel directly into R
Idrees <- read.delim("clipboard")

# View the full dataset in a spreadsheet-style window
View(Idrees)

# Display first and last few rows to verify data structure
head(Idrees)
tail(Idrees)


# 🔄 Step 2: Recode Numeric Demographic Variables
# ------------------------------------------------
# Convert Age from numeric into two age categories
Idrees$Age <- cut(Idrees$Age,
                  breaks = c(19, 21, 24),
                  labels = c("Young", "Older"),
                  include.lowest = TRUE)

# Convert Gender codes to labels
Idrees$Gender <- factor(Idrees$Gender, levels = 1:2, labels = c("Male", "Female"))

# Convert Education codes to labels
Idrees$Edu <- factor(Idrees$Edu, levels = 1:2, labels = c("Graduate", "Undergraduate"))

# Convert Socioeconomic status codes to labels
Idrees$SES <- factor(Idrees$SES, levels = 1:3, labels = c("Low", "Middle", "High"))


# 📊 Step 3: Frequency Tables (SPSS-style)
# ------------------------------------------------
freq_table_gender <- as.data.frame(table(Idrees$Gender))
freq_table_gender$Percent <- round(100 * freq_table_gender$Freq / sum(freq_table_gender$Freq), 1)
freq_table_gender$Cumulative <- cumsum(freq_table_gender$Percent)
colnames(freq_table_gender) <- c("Category", "Frequency", "Percent", "CumulativePercent")
freq_table_gender$Label <- c("Valid", "")

freq_table_age <- as.data.frame(table(Idrees$Age))
freq_table_age$Percent <- round(100 * freq_table_age$Freq / sum(freq_table_age$Freq), 1)
freq_table_age$Cumulative <- cumsum(freq_table_age$Percent)
colnames(freq_table_age) <- c("Category", "Frequency", "Percent", "CumulativePercent")
freq_table_age$Label <- c("Valid", "")

freq_table_edu <- as.data.frame(table(Idrees$Edu))
freq_table_edu$Percent <- round(100 * freq_table_edu$Freq / sum(freq_table_edu$Freq), 1)
freq_table_edu$Cumulative <- cumsum(freq_table_edu$Percent)
colnames(freq_table_edu) <- c("Category", "Frequency", "Percent", "CumulativePercent")
freq_table_edu$Label <- c("Valid", "")

freq_table_ses <- as.data.frame(table(Idrees$SES))
freq_table_ses$Percent <- round(100 * freq_table_ses$Freq / sum(freq_table_ses$Freq), 1)
freq_table_ses$Cumulative <- cumsum(freq_table_ses$Percent)
colnames(freq_table_ses) <- c("Category", "Frequency", "Percent", "CumulativePercent")
freq_table_ses$Label <- c("Valid", "", "")

# Helper function to reorder columns for APA format
reorder_cols <- function(df) {
  df[, c("Label", "Category", "Frequency", "Percent", "CumulativePercent")]
}

freq_table2 <- reorder_cols(freq_table_gender)
freq_table3 <- reorder_cols(freq_table_age)
freq_table4 <- reorder_cols(freq_table_edu)
freq_table5 <- reorder_cols(freq_table_ses)


# 📐 Step 4: Create APA-style Table
# ------------------------------------------------
apa_table_spss <- function(df) {
  border_line <- fp_border(width = 1)
  flextable(df) %>%
    set_header_labels(
      Label = "", 
      Category = "", 
      Frequency = "Frequency", 
      Percent = "Percent", 
      CumulativePercent = "Cumulative Percent"
    ) %>%
    theme_booktabs() %>%
    border_remove() %>%
    border(part = "header", border.top = border_line) %>%
    border(part = "header", border.bottom = border_line) %>%
    border(part = "body", i = nrow(df), border.bottom = border_line) %>%
    fontsize(size = 12, part = "all") %>%
    font(fontname = "Times New Roman", part = "all") %>%
    bold(part = "all", bold = FALSE) %>%
    align(align = "center", part = "all") %>%
    autofit()
}


# 📄 Step 5: Generate Interpretation Text for Each Table
# ------------------------------------------------
interpret_text <- function(table_number) {
  if (table_number == 1) {
    "The sample consisted equally of males and females (50% each), indicating a balanced gender distribution."
  } else if (table_number == 2) {
    "Participants were evenly split between 'Young' and 'Older' age groups, each making up 50% of the sample."
  } else if (table_number == 3) {
    "A majority of the participants were graduates (60%), indicating a slightly higher education level among respondents."
  } else if (table_number == 4) {
    "Most participants belonged to the middle socio-economic status group (40%), suggesting a relatively diverse economic background."
  }
}


# 📤 Step 6: Create Word Document with Tables + Interpretation
# ------------------------------------------------
add_table_with_interpretation <- function(doc, table_title, subtitle, table_obj, table_number) {
  title_fpar <- fpar(
    ftext(table_title, prop = fp_text(font.size = 12, bold = TRUE, font.family = "Times New Roman"))
  )
  subtitle_fpar <- fpar(
    ftext(subtitle, prop = fp_text(font.size = 12, italic = TRUE, font.family = "Times New Roman"))
  )
  
  doc %>%
    body_add_fpar(title_fpar) %>%
    body_add_fpar(subtitle_fpar) %>%
    body_add_flextable(apa_table_spss(table_obj)) %>%
    body_add_par(interpret_text(table_number), style = "Normal") %>%
    body_add_par("")
}

# 📑 Step 7: Compile and Save Document
# ------------------------------------------------
sect_properties <- prop_section(
  page_size = page_size(orient = "portrait", width = 8.27, height = 11.69),
  page_margins = page_mar(top = 1, bottom = 1, right = 1, left = 1.5)
)

doc <- read_docx() %>%
  body_add_par("", style = "Normal") %>%
  body_set_default_section(sect_properties) %>%
  add_table_with_interpretation("Table 4.1 Gender of the respondents", "Gender", freq_table2, 1) %>%
  add_table_with_interpretation("Table 4.2 Age of the respondents", "Age", freq_table3, 2) %>%
  add_table_with_interpretation("Table 4.3 Education of the respondents", "Education", freq_table4, 3) %>%
  add_table_with_interpretation("Table 4.4 Socio-economic status of the respondents", "Socio-economic Status", freq_table5, 4)

dir.create("D:/APA_Tables", showWarnings = FALSE)
print(doc, target = "D:/APA_Tables/APA_Tables_with_Interpretation.docx")
