#Getting Libraries
library(readr)
library(dplyr)
library(stringr)
library(data.table)
library(readxl)
library(officer)

#Reading EMPENG Reports (First set the working environment)
path <- "C:/Users/18360039068941666201/OneDrive - HelloFresh Group/Desktop/Uni_Cansu/AI-SciPro/Employee Engagement/Employee_Engagement_Papers"
setwd(path)
file_vector <- list.files()
docx_list <- file_vector[grepl(".docx", file_vector)]
empeng_df <- data.frame(file_name = character(), text = character(), stringsAsFactors = FALSE)

# Loop through the docx_list
for (i in seq_along(docx_list)) {
  print(i)
  doc <- read_docx(docx_list[i]) 
  document_text <- docx_summary(doc)$text
  document <- data.frame(
    file_name = gsub(x = docx_list[i], pattern = ".docx", replacement = ""),
    text = document_text,
    stringsAsFactors = FALSE
  )
  empeng_df <- rbind(empeng_df, document)
}
str(empeng_df)

# Convert empeng_df to a data.table
empeng_dt <- as.data.table(empeng_df)
empeng_dt <- empeng_dt[, lapply(.SD, paste0, collapse=" "), by = .(file_name)]
empeng_dt[1,]
report_details <- read_excel("C:/Users/18360039068941666201/OneDrive - HelloFresh Group/Report_details.xlsx")

#Join Report Details with empeng_dt
empeng_reports=merge(empeng_dt, report_details, by="file_name")

# Load libraries
library(tm)
library(textclean)
library(wordcloud)
library(ggplot2)
library(viridisLite)
library(RColorBrewer)
library(textstem)

corpus_cleaning <- function(corpus) { 
  corpus <- tm_map(corpus, content_transformer(tolower))
  corpus <- tm_map(corpus, removeNumbers)
  corpus <- tm_map(corpus, removeWords, c(stopwords("en"), "employee", "engagement"))
  corpus <- tm_map(corpus, removePunctuation)
  corpus <- tm_map(corpus, stripWhitespace)
  corpus <- tm_map(corpus, content_transformer(lemmatize_strings))
  return(corpus)
} 

colnames(empeng_reports)[which(names(empeng_reports) == 'file_name')] <- 'doc_id'
empeng_corpus <- VCorpus(DataframeSource(empeng_reports))
empeng_corpus_clean <- corpus_cleaning(empeng_corpus)
empeng_dtm <- DocumentTermMatrix(empeng_corpus_clean)
empeng_dtm_m <- as.matrix(empeng_dtm)


freq <- colSums(empeng_dtm_m)
freq <- sort(freq, decreasing = TRUE)
freq_words <- names(freq)
common_words <- c("common_word1", "common_word2")  # Replace with actual words to be removed
freq_filtered <- freq[!(names(freq) %in% common_words)]

png(filename = "wordcloud_output.png", width = 1200, height = 800, units = "px", res = 100)

wordcloud(words = names(freq_filtered), freq = freq_filtered, max.words = 30, 
          colors = viridis(n = 5), random.color = FALSE)

dev.off()
getwd()

#TOKENIZATION 
install.packages("tokenizers")
library(tokenizers)
# Define a custom tokenizer function for n-gram tokenization
ngram_tokenizer <- function(x) {
  ngram_list <- ngrams(words(x), 2)
  unlist(lapply(ngram_list, paste, collapse = " "))
}

#Tokenize using the custom tokenizer with the cleaned corpus
empeng_tokenize_dtm <- DocumentTermMatrix(empeng_corpus_clean, control = list(tokenize = ngram_tokenizer))
empeng_tokenize_dtm_m <- as.matrix(empeng_tokenize_dtm)
freq_tokenize <- colSums(empeng_tokenize_dtm_m)
freq_tokenize <- sort(freq_tokenize, decreasing = TRUE)
freq_tokenize_words <- names(freq_tokenize)

# Display the first few terms
head(freq_tokenize_words)

wordcloud(freq_tokenize_words, freq_tokenize, max.words = 30, colors = viridis(n = 5), random.color = FALSE)
png("wordcloud2.png", width = 1200, height = 800, units = "px", res = 100)
dev.off()
getwd()

library(syuzhet)
library(tm)
library(SnowballC)
library(stringr)
library(textclean)
library(wordcloud)
library(ggplot2)
library(viridisLite)
library(tokenizers)
library(fmsb)
library(tidytext)
library(dplyr)
library(tidyr)
library(reshape2)
library(radarchart)
library(plotly)

corpus_cleaning <- function(corpus) {
corpus <- tm_map(corpus, content_transformer(tolower))
corpus <- tm_map(corpus, removeNumbers)
corpus <- tm_map(corpus, removeWords, stopwords("en"))
corpus <- tm_map(corpus, content_transformer(iconv), "latin1", "ASCII", sub=" ")
corpus <- tm_map(corpus, content_transformer(str_replace_all), "[[:punct:]]", "")
corpus <- tm_map(corpus, content_transformer(str_replace_all), "[^[:alnum:]]", " ")
corpus <- tm_map(corpus, content_transformer(str_replace_all), '\\w{20,}', " ")
corpus <- tm_map(corpus, stripWhitespace)
corpus <- tm_map(corpus, stemDocument)
corpus <- tm_map(corpus, PlainTextDocument)
return(corpus)
}
empeng_corpus <- VCorpus(DataframeSource(empeng_reports))
meta(empeng_corpus, tag = "id") <- empeng_reports$file_id  
all_doc_ids <- sapply(empeng_corpus, function(doc) meta(doc, "id"))
print(all_doc_ids)
empeng_corpus_clean <- corpus_cleaning(empeng_corpus)

empeng_dtm <- DocumentTermMatrix(empeng_corpus_clean, control = list(weighting = weightTfIdf, removeSparseTerms = 0.999))
dist_matrix <- proxy::dist(as.matrix(empeng_dtm), method = "canberra")
clustering.hierarchical <- hclust(dist_matrix, method = "ward.D2")
clusters <- cutree(clustering.hierarchical, k = 10)

cluster_documents <- lapply(1:10, function(k) {
  corpus_subset <- empeng_corpus_clean[clusters == k]
  return(corpus_subset)
})

perform_sentiment_analysis <- function(corpus_subset) {
  dtm <- DocumentTermMatrix(corpus_subset)
  df <- tidy(dtm)
  doc_ids <- sapply(corpus_subset, meta, "id")
  doc_term_counts <- rowSums(as.matrix(dtm) > 0) 
  df$document <- rep(doc_ids, times = doc_term_counts)
  
 
  sentiment_df <- df %>%
    inner_join(get_sentiments("nrc"), by = c("term" = "word"), relationship = "many-to-many") %>%
    filter(!sentiment %in% c("positive", "negative")) %>% 
    group_by(document) %>%
    count(sentiment) %>%
    ungroup() %>%
    spread(key = sentiment, value = n, fill = 0)
  
  return(sentiment_df)
}
cluster_sentiments <- lapply(cluster_documents, perform_sentiment_analysis)

plot_radar_chart <- function(sentiment_df, cluster_number) {
  sentiments <- c("anger", "anticipation", "disgust", "fear", "joy", "sadness", "surprise", "trust")
  scores <- sapply(sentiments, function(s) {
    if(s %in% names(sentiment_df)) {
      return(sentiment_df[[s]])
    } else {
      return(0)
    }
  }, USE.NAMES = FALSE)
  
  # Normalize sentiment scores
  max_score <- max(scores)
  if (max_score > 0) {
    normalized_scores <- scores / max_score
  } else {
    normalized_scores <- scores 
  }
  
  radar_data <- data.frame(
    sentiment = factor(c(sentiments, sentiments[1]), levels = sentiments),
    score = c(normalized_scores, normalized_scores[1]),
    cluster = factor(rep(cluster_number, length(sentiments) + 1))
  )

  radarchart <- ggplot(radar_data, aes(x = sentiment, y = score, group = cluster)) +
    geom_line(aes(color = cluster), size = 0.75) +
    geom_point(aes(color = cluster), size = 2) +
    theme_light() +
    ggtitle(paste("Cluster", cluster_number, "Sentiment Analysis")) +
    xlab("") +
    ylab("Normalized Score") +
    scale_color_manual(values = c("#CC6666", "#9999CC", "#66CC99", "#CC9966")) +
    coord_polar() 

  print(radarchart)
  
  
  file_name <- paste0("cluster_", cluster_number, "_radar_chart.png")
  ggsave(file_name, plot = radarchart, width = 10, height = 8, dpi = 300)
}

lapply(seq_along(cluster_sentiments), function(i) {
  plot_radar_chart(cluster_sentiments[[i]], i)
})
