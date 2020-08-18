library(forestplot)
library(xlsx)

datadir = 'data'
file = list.files(datadir)[endsWith(list.files(datadir),'xlsx')]
filelist = file[c(-2,-8)]#1548 and 65682 are 2 datasets from other platforms


data = list()
for (i in 1:length(filelist)){
  data[[i]]<-read.xlsx(paste(datadir,'/',filelist[[i]],sep = ''),header = T,sheetIndex = 1)[,-1]
}
data2 = list()
data2[[1]] <- read.xlsx(paste(datadir,'/',file[[2]],sep = ''),header = T,sheetIndex = 1)[,-1]
data2[[2]] <- read.xlsx(paste(datadir,'/',file[[8]],sep = ''),header = T,sheetIndex = 1)[,-1]

filename <- gsub("\\.xlsx",'',file)

filename1<-paste('GSE',filename[c(1,3:7,9:11)],sep = '')
filename2 <- c(paste('E-MTAB-',filename[2],sep =''),paste('GSE',filename[8],sep = ''))


dataall <- rbind(data[[1]],data[[2]],data[[3]],data[[4]],data[[5]],data[[6]],data[[7]],data[[8]],
                 data[[9]],data2[[1]],data2[[2]])

data2[[3]] = dataall

out <- list()

for (i in 1:6){
  mean = c()
  upper = c()
  lower = c()
  p = c()
  foldc = c()
  for (j in 1:9){
    t = t.test(data[[j]][,i][data[[j]]$label == 1],data[[j]][,i][data[[j]]$label == 0])
    mean = c(mean, as.numeric(t$estimate[1] - t$estimate[2]))
    upper = c(upper,t$conf.int[2])
    lower = c(lower,t$conf.int[1])
    p = c(p,t$p.value)
  }
  for (j in 1:3){
    t = t.test(data2[[j]][,i][data2[[j]]$label == 1],data2[[j]][,i][data2[[j]]$label == 0])
    mean = c(mean, as.numeric(t$estimate[1] - t$estimate[2]))
    upper = c(upper,t$conf.int[2])
    lower = c(lower,t$conf.int[1])
    p = c(p,t$p.value)
  }
  
  name = c(filename1,filename2,'Summary')
  temp <- data.frame(name = name,mean = mean,lower = lower, upper = upper,p_value = p)
  out[[i]] = temp
}


# dev.off()
par(mfrow = c(1,4))

mean = as.matrix(cbind(out[[1]]$mean,out[[2]]$mean,out[[3]]$mean,out[[4]]$mean))
lower = as.matrix(cbind(out[[1]]$lower,out[[2]]$lower,out[[3]]$lower,out[[4]]$lower))
upper = as.matrix(cbind(out[[1]]$upper,out[[2]]$upper,out[[3]]$upper,out[[4]]$upper))
forestplot(list(out[[1]]$name),mean,lower,upper,grid = T)

dir = 'figure/'
png(paste(dir,colnames(data[[1]])[1],'.png',sep = ''),bg = 'transparent',height = 400,width = 450)
forestplot(list(rep('',12)),out[[1]]$mean,out[[1]]$lower,out[[1]]$upper,
           title = colnames(data[[1]])[1],grid = T,
           xlab = 'Standardized Difference Mean',
           col = fpColors(box = "#1950A0",lines = "#9DB7DE",zero = 'grey50',summary = 'brown1'),
           cex.lab = 3,
           xticks = 0:5,
           txt_gp = fpTxtGp(xlab = gpar(cex = 1.5),cex = 1.8),
           lineheight = 'auto',
           colgap = unit(6,'mm'),
           lwd.ci = 2,
           ci.vertices = T,
           ci.vertices.height = 0.2,is.summary = c(rep(F,11),T)
           )
dev.off()
png(paste(dir,colnames(data[[1]])[2],'.png',sep = ''),bg = 'transparent',height = 400,width = 600)
forestplot(list(out[[2]]$name),out[[2]]$mean,out[[2]]$lower,out[[2]]$upper,
           title = colnames(data[[1]])[2],grid = T,
           xlab = 'Standardized Difference Mean',
           col = fpColors(box = "#1950A0",lines = "#9DB7DE",zero = 'grey50',summary = 'brown1'),
           cex.lab = 3,
           xticks = c(0:5),
           lineheight = 'auto',
           txt_gp = fpTxtGp(xlab = gpar(cex = 1.5),cex = 1.8),
           colgap = unit(6,'mm'),
           lwd.ci = 2,
           ci.vertices = T,
           ci.vertices.height = 0.2,is.summary = c(rep(F,11),T)
)
dev.off()
png(paste(dir,colnames(data[[1]])[3],'.png',sep = ''),bg = 'transparent',height = 400,width = 700)
forestplot(list(rep('',12)),out[[3]]$mean,out[[3]]$lower,out[[3]]$upper,
           title = colnames(data[[1]])[3],
           grid = T,
           xlab = 'Standardized Difference Mean',
           col = fpColors(box = "#1950A0",lines = "#9DB7DE",zero = 'grey50',summary = 'brown1'),
           txt_gp = fpTxtGp(xlab = gpar(cex = 1.5),cex = 1.8),
           lineheight = 'auto',
           colgap = unit(6,'mm'),
           lwd.ci = 2,
           xticks = 0:8,
           ci.vertices = T,
           ci.vertices.height = 0.2,is.summary = c(rep(F,11),T))
dev.off()
png(paste(dir,colnames(data[[1]])[4],'.png',sep = ''),bg = 'transparent',height = 400,width = 450)
forestplot(list(rep('',12)),out[[4]]$mean,out[[4]]$lower,out[[4]]$upper,
           title = colnames(data[[1]])[4],
           grid = T,
           xlab = 'Standardized Difference Mean',
           col = fpColors(box = "#1950A0",lines = "#9DB7DE",zero = 'grey50',summary = 'brown1'),
           cex.lab = 3,
           lineheight = 'auto',
           colgap = unit(6,'mm'),
           xticks = 0:5,
           lwd.ci = 2,
           txt_gp = fpTxtGp(xlab = gpar(cex = 1.5),cex = 1.8),
           ci.vertices = T,
           ci.vertices.height = 0.2,
           is.summary = c(rep(F,11),T))
dev.off()

png(paste(dir,colnames(data[[1]])[5],'.png',sep = ''),bg = 'transparent',height = 400,width = 800)
forestplot(list(out[[5]]$name),out[[5]]$mean,out[[5]]$lower,out[[5]]$upper,
           title = colnames(data[[1]])[5],
           grid = T,
           xlab = 'Standardized Difference Mean',
           col = fpColors(box = "#1950A0",lines = "#9DB7DE",zero = 'grey50',summary = 'brown1'),
           cex.lab = 3,
           lineheight = 'auto',
           xticks = 0:8,
           colgap = unit(6,'mm'),
           txt_gp = fpTxtGp(xlab = gpar(cex = 1.5),cex = 1.8),
           lwd.ci = 2,
           ci.vertices = T,
           ci.vertices.height = 0.2,
           is.summary = c(rep(F,11),T))
dev.off()

png(paste(dir,colnames(data[[1]])[6],'.png',sep = ''),bg = 'transparent',height = 400,width = 1500)
forestplot(list(out[[6]]$name),out[[6]]$mean,out[[6]]$lower,out[[6]]$upper,
           title = 'LIFTS',
           grid = T,
           xlab = 'Standardized Difference Mean',
           col = fpColors(box = "#1950A0",lines = "#9DB7DE",zero = 'grey50',summary = 'brown1'),
           cex.lab = 3,
           lineheight = 'auto',
           colgap = unit(6,'mm'),
           txt_gp = fpTxtGp(xlab = gpar(cex = 1.5),cex = 1.8),
           lwd.ci = 2,
           ci.vertices = T,
           cex.axis = 2,
           xticks = 0:16,
           ci.vertices.height = 0.2,
           is.summary = c(rep(F,11),T))
dev.off()

