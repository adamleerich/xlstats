set.seed(12)
x <- rlnorm(1000, 13, 1.5)

clipr::write_clip(x, 'table')

fitdistrplus::fitdist(x, 'lnorm', 'mme')
