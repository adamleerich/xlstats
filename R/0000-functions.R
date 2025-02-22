Int32 <- function(x) {
  # Mimics C's two's complement overflow of signed Int32
  x <- sign(x)*floor(abs(x))
  x <- x %% 2^32
  ifelse(x >= 2^31, x - 2^32, x)
}

