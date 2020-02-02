library(Cairo)

CairoPDF(width = 7, height = 7)
plot(x = iris$Sepal.Length, y = iris$Sepal.Width,
     pch = c(21, 24, 25)[unclass(iris$Species)],
     bg = c("red", "green3", "blue")[unclass(iris$Species)],
     las = 1, panel.first = grid(), tcl = 0.25, cex = 2,
     main = "iris plot sample")
legend("topright", c("setosa", "versicolor", "virginica"), pch = c(21, 24, 25),
       bg = "white", pt.bg = c("red", "green3", "blue"), cex = 1)
dev.off()