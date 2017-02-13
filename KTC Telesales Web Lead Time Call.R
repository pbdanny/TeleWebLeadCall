# Read data -----
# library(readxl)
# web <- read_excel(path = "/Users/Danny/Documents/R Project/KTC.Tele.Web.Time/List Web DateTime.xlsx", 
#                  sheet = 1, col_types = rep("text",7))
# colnames(web) <- make.names(colnames(web))

# save(web, file = "TeleWebCallTime.RData")
load(file = "TeleWebCallTime.RData")

# Manipulate date time data ----

# Date data in Excel stored in format no. of day from origin (1899-12-30)
# While POSIXct use format no. of second from origin
# Then convert form #day -> #second 

textWinDateConv <- function (d) {
  t <- as.POSIXct(as.numeric(d) * 60*60*24 , # Convert days to second for POSIXct time scales 
                  # Set origin to Bankok minus 17 min 56 sec for correct timezone orgin before 1920
                  # https://en.wikipedia.org/wiki/Time_in_Thailand
                  origin = (as.POSIXct("1899-12-30", tz = "Asia/Bangkok") - (17*60+56)),
                  tz = "Asia/Bangkok") # To assure out put time zone
  return(t)
}

# Combine date & time called
web$DateTimeCall <- as.numeric(web$DateCall) + as.numeric(web$TimeCall)
web$DateCall <- NULL
web$TimeCall <- NULL

# Convert all Excel date-time to POSIXct
web$DateTimeSubmit <- textWinDateConv(web$DateTimeApply)
web$DateTimeExport <- textWinDateConv(web$DateTimeExport)
web$DateTimeCall <- textWinDateConv(web$DateTimeCall)

# Change incorrect year (1959 to 2016, +57 yr) - DateTimeApply and DateTimeExport
web$DateTimeSubmit <- as.POSIXlt(web$DateTimeSubmit)
web$DateTimeSubmit$year <- web$DateTimeSubmit$year + 57
web$DateTimeSubmit <- as.POSIXct(web$DateTimeSubmit)

web$DateTimeExport <- as.POSIXlt(web$DateTimeExport)
web$DateTimeExport$year <- web$DateTimeExport$year + 57
web$DateTimeExport <- as.POSIXct(web$DateTimeExport)

# Function find next working day

load(file = "Thai Bank Holiday.RData")

nextWorkDay <- function (d, EndHr, Holiday) {
  out <- list()
  for (i in 1:length(d)) {
    while (as.integer(format(d[i], "%H")) >= EndHr | (format(d[i], "%w") %in% c("0", "6")) | 
           (format(d[i], "%Y-%m-%d") %in% Holiday)) { 
      d[i] <- d[i] + (24*60*60)  # Add 1 day (= add 24*60*60 seconds)
      d[i] <- as.POSIXct(paste0(format(d[i], "%Y-%m-%d"), " 08:00:00"))  # Move to start working hours
    }
    out <- rbind(out, format(d[i], "%Y-%m-%d %H:%M:%S"))
  }
  return(unlist(out))
}

# Call function nextWorkDay
web$DateTimeStart <- nextWorkDay(web$DateTimeSubmit, 16, thaiBankHoliday)

# Convert text to POSIXct
web$DateTimeStart <- as.POSIXct(web$DateTimeStart, 
                              origin = (as.POSIXct("1899-12-30", tz = "Asia/Bangkok") - (17*60+56)),
                              tz = "Asia/Bangkok")
# Data Audit -----
web$hrs.diff <- as.numeric(difftime(web$DateTimeCall, web$DateTimeStart, units ="hours"))

# Correct hrs.diff < 0 caused by DateTimeCall before DateTimeStart caused by leadSubmit after 16:00
# Then change DateTimeStart back to DateTimeSubmit for hrs.diff < 0
rep.idx <- which(web$hrs.diff < 0)
web[rep.idx, "DateTimeStart"] <- web[rep.idx, "DateTimeSubmit"]
web$hrs.diff <- as.numeric(difftime(web$DateTimeCall, web$DateTimeStart, units ="hours"))

# Select only feasible data
df <- web[web$hrs.diff > 0 & !is.na(web$hrs.diff) & web$hrs.diff < 168, ]  # Choose hrs.diff < 168 hr(=7 days)
rm(rep.idx)

# Export Data for error check
df.err <- web[!(web$Identifier.ID %in% df$Identifier.ID), ]
library(xlsx)
write.xlsx2(df.err$Identifier.ID, file = "Error.xlsx", row.names = FALSE)
rm(df.err)

# Exploratory Data Analysis ----
# save(df, file = "Tele Web Call Time-Audited.RData")

load(file = "Tele Web Call Time-Audited.RData")

# Basict Stat Part

statsummary.df <- function(df) {
  sample.mean <- mean(df$hrs.diff)
  sample.sd <- sd(df$hrs.diff)
  sample.median <- median(df$hrs.diff)
  n <- nrow(df)
  err <- qt(0.975, df = n - 1)*sample.sd/sqrt(n)
  lower <- sample.mean - err
  upper <- sample.mean + err
  out <- data.frame(c("Sample.mean" = sample.mean,
                      "Sample.SD" = sample.sd,
                      "95% lower conf. interval" = lower,
                      "95% upper conf. interval" = upper,
                      "Median" = sample.median))
  colnames(out) <- "Hrs.diff"
  return(out)
}

library(ggplot2)
g <- ggplot(df, aes(x = hrs.diff)) + 
  geom_histogram(binwidth = 1, fill = "green", color = "green", alpha = .5) +
  coord_cartesian(xlim = c(0, 50)) +
  geom_vline(xintercept = 2, color = "red") +
  geom_vline(xintercept = 5, color = "green", linetype = 2) +
  ggtitle("Total Hrs. Diff : Time Call - Time Apply") +
  xlab("Time Call - Time Apply (hrs)")
g

# add flag holiday
load(file = "Thai Bank Holiday.RData")
df$submitHoliday <- ifelse(format(df$DateTimeSubmit, "%Y-%m-%d") %in% thaiBankHoliday |
                             format(df$DateTimeSubmit, "%w") %in% c("0", "6"), "Holiday", "Not Holiday")
# Plot Holiday

g <- ggplot(df, aes(x = hrs.diff, fill = as.factor(submitHoliday), color = as.factor(submitHoliday))) + 
  geom_histogram(binwidth = 1, alpha = .5, position = "identity") +
  coord_cartesian(xlim = c(0, 50)) +
  ggtitle("Hrs Diff : Holiday vs. Not Holiday") +
  xlab("Time Call - Time Apply (hrs)")
g

stat.holi <- statsummary.df(df[df$submitHoliday == "Holiday", ])
stat.notholi <- statsummary.df(df[df$submitHoliday == "Not Holiday",])
t.test(hrs.diff ~ submitHoliday, data = df)
# T-test show diffrence mean lead submit holiday ~ holiday

df$submitDayofWeek <- as.integer(format(df$DateTimeSubmit, "%w"))
df$DayofWeek<- factor(df$submitDayofWeek, 
                          labels = c("Sun", "Mon", "Tue", "Wed", "Thr", "Fri", "Sat"))

i <- ggplot(df, aes(x = hrs.diff, color = DayofWeek)) + 
  geom_density(alpha = .25) +
  coord_cartesian(xlim = c(0, 50)) +
  geom_vline(xintercept = 2, color = "red")
i
# Geom desity show lead Sat & Sun has higher means

# Export data for recheck
write.csv(df, file = "write.csv")