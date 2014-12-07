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

tabRefName = "Calibration regressions"
tab_indexStart <- which(wb_sheets == tabRefName) +1

# extract a sheet into a data frame.
#dat <- readWorksheet(wbObject, sheet = getSheets(wbObject))
tab_indexList = tab_indexStart:length(wb_sheets)

for (tab_index in tab_indexList) {
  cat(paste("..", (tab_index - tab_indexStart)+1 , ".."))
  tabDat <- readWorksheet(wb_object, sheet = wb_sheets[tab_index], header = TRUE)
  
  plot(tabDat$Time..sec., tabDat$BS.Ca43, type="l", lwd=4)
  title(wb_sheets[tab_index])
  xy <- identify(tabDat$Time..sec., tabDat$BS.Ca43)
  # lines(tabDat$Time..sec.[xy[1]:xy[2]],tabDat$BS.Ca43[xy[1]:xy[2]],col="red",lwd=6)

  # chop the columns that contain 'mol' (that's the grep bit) based on user input (that's th exy stuff) indices.

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
  savename = gsub(" ", "", paste(filename,"_TRIMMED_",wb_sheets[tab_index],".csv"))
  write.csv(subTabDat_molCols, file = savename,row.names=FALSE)
}


