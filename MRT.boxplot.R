#create multivariate regression tree with multiobjective boxplots at each leaf
#assumes input data is an MOEA tradeoff set in an Excel file with the following format:
#rows are portfolios
#columns in the following order: Archive ID, decision levers, objectives

#set directory to location of MOEA archive
setwd("C:/Users/rebeccasmith/Desktop/MRT")

workbook.name = "Archive.Eldorado.4C.xlsx"

num.lev = 13
num.obj = 7

#packages etc.
options( java.parameters = "-Xmx4g" )
require(XLConnect)
require(mvpart)
require(TeachingDemos)


wb = loadWorkbook(workbook.name)
df.id = readWorksheet(wb, sheet = "Sheet1", header = TRUE)

#remove ID col
df = df.id[,-1]

#rescale objective performance values to [0,1]
myrescale <- function(x) (x-min(x))/(max(x) - min(x))

scaled.obj = apply(df[,(num.lev+1):ncol(df)], 2, myrescale)

scaled.df = cbind(scaled.obj, df[,1:num.lev])

myform = (data.matrix(scaled.df[,1:num.obj])) ~ (WS_Shares + Ag3_Rights
                                           + Cons_Factor + Dist_Eff + Indust_Rights
                                           + Vol_Res2 + Vol_WestRes + Vol_Xres + GP
                                           + Int_Shares + Ag2_Shares + Exchange + Lease_Ag2Res)

#create MRT and plot tree performance vs. tree size to determine best tree size
#set very small complexity parameter to see how error evolves with tree size
stop.rule = 0.001

mrt.error = mvpart(myform, data = scaled.df, pretty = T, xv = "min", minauto = F,
                 which = 4, bord = T, uniform = F, text.add = T, branch = 1,
                 xadj = .7, yadj = 1.2, use.n = T, margin = 0.05, keep.y = F,
                 bars = F, all.leaves = F, control = rpart.control(cp = stop.rule, xval = 10), #10-fold cross validation
                 plot.add = F)

#comment svg and dev.off to show tree in plot window instead of saving to file
svg("Eldorado.4C.MRT.error.svg", width = 8, height = 6, pointsize = 9)
plotcp(mrt.error, upper = "size") # size is for # leaves, splits is for # splits
points(mrt.error$cptable[,"rel error"], col = "red")
legend(x = "topright", legend = c("rel error", "xvalidated rel error"), 
       col = c("red", "blue"), pch = c(1, 1))
dev.off()

full.table = as.data.frame(mrt.error$cptable)
print(full.table)


#after examining error plot, specify tree either by number of leaves or complexity parameter
#modify revised.stop variable depending on method of specifying tree
num.leaves = 14
final.cp = 0.01

revised.stop = full.table$CP[full.table$nsplit == (num.leaves-1)]

#create a tree to get coordinates for plotting boxplots later
pre.plot = mvpart(myform, data = scaled.df, pretty = T, xv = "min", minauto = F,
                  which = 4, bord = T, uniform = F, text.add = T, branch = 1,
                  xadj = .7, yadj = 1.2, use.n = T, margin = 0.05, keep.y = F,
                  bars = F, all.leaves = F, control = rpart.control(cp = revised.stop, xval = 10),
                  plot.add = F)

getxy = as.data.frame(cbind(plot(pre.plot, uniform = F)$x, plot(pre.plot, uniform = F)$y))
dev.off()

#create final tree with boxplot leaves
#comment svg and dev.off to show tree in plot window instead of saving to file

svg("Eldorado.4C.MRT.svg", width = 18, height = 7.5, pointsize = 9) #change width and height depending on tree size

final.mrt = mvpart(myform, data = scaled.df, pretty = T, xv = "min", minauto = T,
                 which = 4, bord = T, uniform = F, text.add = T, branch = 1,
                 xadj = .7, yadj = 1.2, use.n = T, margin = 0.05, keep.y = F,
                 bars = F, all.leaves = F, control = rpart.control(cp = revised.stop, xval = 10),
                 plot.add = T)

title(main = paste("Eldorado Utility 1C Tradeoffs"), cex = 2)
legend("topleft",legend = colnames(scaled.df[,1:num.obj]), 
       fill = c("blue", "grey","red", "purple","yellow", "green","orange"), #add or remove colors depending on num.obj
       bty = "n")

#add column to df that records which leave each portfolio belongs in
leaf.assign = cbind(as.matrix(final.mrt$where), scaled.df) 
colnames(leaf.assign) = c("Leaf", colnames(scaled.df))

#specify coordinates at which leaves will be plotted
leaves = as.numeric(levels(as.factor(pre.plot$where))) 

#add boxplots to terminal nodes
for (i in 1:length(leaves)){
  subplot(boxplot.matrix(as.matrix(scaled.df[which(leaf.assign$Leaf==leaves[i]),1:7]), #when specifying which columns, must use numeric values (so fill in 1:#objectives)
                         col = c("blue", "grey","red", "purple","yellow", "green","orange"), #add or remove colors depending on num.obj
                         xaxt = "n", yaxt = "n", xlab = paste("Leaf",i),
                         cex = .5, labels = F, bg = "white"),
          
          x = .95*getxy$V1[leaves[i]], #may need to adjust x & y positions to prevent boxplots from crowding
          y = .92*getxy$V2[leaves[i]],
          size = c(0.85, .8))
}

dev.off()

pruned.table = as.data.frame(final.mrt$cptable)
print(pruned.table)
