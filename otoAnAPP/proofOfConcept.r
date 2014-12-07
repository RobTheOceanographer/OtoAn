# Proof of concept for loading, cropping, plotting the otolith mid-analysis datasheet.

# Setup the libraries we need to access the excel dataset.
# install.packages("XLConnect")
options(java.parameters = "-Xmx4g" ) # gets around java memory limits.
# options(java.parameters = "-Xmx1600")
library("XLConnect")

filename <- "/home/dev/projects/otoAn/exampleInputData/BLOCK 33.xls"
# Read the workbook
message("***** Loading in the data. This could take a while! *****")
wb_object<-loadWorkbook(filename)

# Save each sheet's name as a vector
wb_sheets <-getSheets(wb_object)

tabRefName = "Calibration regressions"
tab_indexStart <- which(wb_sheets == tabRefName) +1

# extract a sheet into a data frame.
#dat <- readWorksheet(wbObject, sheet = getSheets(wbObject))
tab_indexList <- tab_indexStart:length(wb_sheets)

message("LOADING ALL THE SHEETS AFTER THE CALIBRATION WORKSHEET...")
dat<-lapply(seq_along(wb_sheets[tab_indexList]),function(i) readWorksheet(wb_object,sheet=wb_sheets[tab_indexList][i]))


message("***** Data loaded, starting plotting now *****")
savenameIDX = 1
for (tabDat in dat) {
  # cat(paste("..", atab , ".."))
#  tabDat <- readWorksheet(wb_object, sheet = wb_sheets[tab_index], header = TRUE)
  message("***** Select start then end point and click finish *****")
  plot(tabDat$Time..sec., tabDat$BS.Ca43, type="l", lwd=3)
  title(wb_sheets[tab_indexList][savenameIDX])
  xy <- identify(tabDat$Time..sec., tabDat$BS.Ca43)
  dev.off()
  # lines(tabDat$Time..sec.[xy[1]:xy[2]],tabDat$BS.Ca43[xy[1]:xy[2]],col="red",lwd=6)

  # chop the columns that contain 'mol' (that's the grep bit) based on user input (that's th exy stuff) indices.
  message("***** SAVING THE TRIMMED DATA *****")
  subTabDat_molCols <- tabDat[xy[1]:xy[2],grep("mol", colnames(tabDat))]
  subTabDat_molCols$time <- tabDat$Time..sec.[xy[1]:xy[2]]
  
  #remove the dots from the col names.
  names(subTabDat_molCols) <- gsub(x = names(subTabDat_molCols),
     pattern = "\\.",
     replacement = "_")
  # remove the trailing _
  names(subTabDat_molCols) <- gsub("_$","", names(subTabDat_molCols), perl=T)
  # remove double _
  names(subTabDat_molCols) <- gsub("__","_", names(subTabDat_molCols), perl=T)
  
  # plot(subTabDat_molCols$time, subTabDat_molCols$Li7_Ca43_umol_mol,type='l')
  #savename = gsub(" ", "", paste("TRIMMED_", basename(filename), ".csv"))
  savename = gsub(" ", "", paste(filename,"_TRIMMED_",wb_sheets[tab_indexList][savenameIDX],".csv"))
  savenameIDX <-savenameIDX + 1
  write.csv(subTabDat_molCols, file = savename,row.names=FALSE)
  rm(subTabDat_molCols)
}

message("Cleaning up...")
rm(list=ls()) 
message("***** FINISHED *****")

