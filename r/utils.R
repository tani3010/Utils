
setwd.currentFilePath <- function() {
  frame_files <- Filter(Negate(is.null), lapply(sys.frames(), function(x) x$ofile))
  setwd(dirname(frame_files[[length(frame_files)]]))
}