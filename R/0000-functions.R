Int32 <- function(x) {
  # Mimics C's two's complement overflow of signed Int32
  x <- sign(x)*floor(abs(x))
  x <- x %% 2^32
  ifelse(x >= 2^31, x - 2^32, x)
}

# Int32(-12387.5)
# 
# 
# Int32(129347812983)
# Int32(-987917263)
# Int32(19273981729381)
# Int32(-9879879878768)
# Int32(1287)
# Int32(-98798576)
# 
# Int32(-112973323113650)
# 
# 
# -112973323113650 %% 2^32
