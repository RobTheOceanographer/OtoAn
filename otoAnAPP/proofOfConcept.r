# Proof of concept for loading, cropping, plotting the otolith mid-analysis datasheet.

# Setup the libraries we need to access the excel dataset.
# install.packages("XLConnect")
options(java.parameters = "-Xmx4g" ) # gets around java memory limits.
# options(java.parameters = "-Xmx1600")
library("XLConnect")

filename <- "/home/dev/projects/otoAn/exampleInputData/BLOCK 33.xls"
# Read the workbook 
wb_object<-loadWorkbook(filename)

# Save each sheet's name as a vector
wb_sheets <-getSheets(wb_object)

tab_index <- which(wb_sheets == "Calibration regressions")

# extract a sheet into a data frame.
#dat <- readWorksheet(wbObject, sheet = getSheets(wbObject))

dat <- readWorksheet(wb_object, sheet = wb_sheets[10], header = TRUE)


plot(dat$Time..sec., dat$BS.Ca43, type="l", lwd=4)
xy <- identify(dat$Time..sec., dat$BS.Ca43)

lines(dat$Time..sec.[xy[1]:xy[2]],dat$BS.Ca43[xy[1]:xy[2]],col="red",lwd=6)

# install.packages("stringr")
library(stringr)
#str_locate(colnames(dat)[i], "mol") 
d <- lapply(seq_along(colnames(dat)),function(i) str_locate(colnames(dat)[i], "mol")[1])
d1 <- lapply(seq_along(d), function(ii) !is.na(as.numeric(d[ii])))

d2<-matrix(d1)

colnames(dat)
dat[d1]


