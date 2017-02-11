# Read data -----
library(readxl)
web <- read_excel(path = "/Users/Danny/Documents/R Project/KTC.Tele.Web.Time/List Web DateTime.xlsx", 
                  sheet = 1, col_types = rep("text",7))
colnames(web) <- make.names(colnames(web))

save(web, file = "TeleWebCallTime.RData")
load(file = "TeleWebCallTime.RData")

View(head(web))
str(web)
summary(web)
View(web[is.na(web$Identifier.ID), ])

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
web$DateTimeSubmit <- textWinDateConv(web$DateTimeSubmit)
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

load(file = "BankSpecialHoliday.RData")

nextWorkDay <- function (d, EndHr, Holiday) {
  out <- list()
  for (i in 1:length(d)) {
    while (as.integer(format(d[i], "%H")) >= EndHr | format(d[i], "%w") %in% c("0", "6") | 
           format(d[i], "%Y-%m-%d") %in% Holiday) { 
      d[i] <- d[i] + (24*60*60)  # Add 1 day (= add 24*60*60 seconds)
      d[i] <- as.POSIXct(paste0(format(d[i], "%Y-%m-%d"), " 08:00:00"))  # Move to start working hours
    }
    out <- rbind(out, format(d[i], "%Y-%m-%d %H:%M:%S"))
  }
  return(unlist(out))
}

# Call function nextWorkDay
web$LeadLoaded <- nextWorkDay(web$DateTimeSubmit, 16, thaiBankHoliday)

# Convert text to POSIXct
web$LeadLoaded <- as.POSIXct(web$LeadLoaded, 
                              origin = (as.POSIXct("1899-12-30", tz = "Asia/Bangkok") - (17*60+56)),
                              tz = "Asia/Bangkok")
# Data analysis -----
web$hrs.diff <- as.numeric(difftime(web$DateTimeCall, web$LeadLoaded, units ="hours"))

# Plot all time called 
library(ggplot2)
library(plotly)
g <- ggplot(web[web$hrs.diff > 0, ], aes(x = hrs.diff)) + 
  geom_histogram(binwidth = 1, fill = "green", color = "green", alpha = .5) +
  coord_cartesian(xlim = c(0, 50)) +
  geom_vline(xintercept = 2, color = "red")
ggplotly(g)

# add flag holiday
web$submitHoliday <- ifelse(format(web$DateTimeSubmit, "%Y-%m-%d") %in% thaiBankHoliday |
                             format(web$DateTimeSubmit, "%w") %in% c("0", "6"), "Holiday", "Not Holiday")

h <- ggplot(web[web$hrs.diff > 0, ], aes(x = hrs.diff, fill = as.factor(submitHoliday))) + 
  geom_histogram(binwidth = 1, alpha = .5, position = "identity") +
  coord_cartesian(xlim = c(0, 50)) +
  geom_vline(xintercept = 2, color = "red")
ggplotly(h)

t.test(hrs.diff ~ submitHoliday, data = web[web$hrs.diff > 0, ])
# T-test show diffrence mean lead submit holiday ~ holiday

web$submitDayofWeek <- as.integer(format(web$DateTimeSubmit, "%w"))
web$DayofWeek<- factor(web$submitDayofWeek, 
                          labels = c("Sun", "Mon", "Tue", "Wed", "Thr", "Fri", "Sat"))

i <- ggplot(web[web$hrs.diff > 0, ], aes(x = hrs.diff, color = DayofWeek)) + 
  geom_density(alpha = .25) +
  coord_cartesian(xlim = c(0, 50)) +
  geom_vline(xintercept = 2, color = "red")
ggplotly(i)
# Geom desity show lead Sat & Sun has higher means