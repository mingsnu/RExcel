## Start R as administrator:For windows 7 user, you need to right-click the R icon,
## and click run as admistrator. For windows XP user, you should login the system
## as admistrator(if you have only one account in your computer, that account is
## admistrator.
remotefile <- "http://rcom.univie.ac.at/download/current/statconnDCOM.latest.exe"
localfile <- paste(tempdir(),"statconnDCOM.latest.exe", sep = "\\")
download.file(remotefile,localfile,mode="wb")
shell(localfile)
install.packages(c("rscproxy","rcom"))
library(rcom)
comRegisterRegistry()
install.packages("RExcelInstaller")
library(RExcelInstaller)
installRExcel()
install.packages(c("RthroughExcelWorkbooksInstaller","Rcmdr","RcmdrPlugin.HH"))
library(RthroughExcelWorkbooksInstaller)
installRthroughExcel()
#install.packages("Rcmdr",depends=T)