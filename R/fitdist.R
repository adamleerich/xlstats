fitdist <- function (data, distr, method = c("mle", "mme", "qme", "mge", 
    "mse"), start = NULL, fix.arg = NULL, discrete, keepdata = TRUE, 
    keepdata.nb = 100, calcvcov = TRUE, ...) 
{
    if (!is.character(distr)) 
        distname <- substring(as.character(match.call()$distr), 
            2)
    else distname <- distr
    ddistname <- paste("d", distname, sep = "")
    if (!exists(ddistname, mode = "function")) 
        stop(paste("The ", ddistname, " function must be defined"))
    if (missing(discrete)) {
        if (is.element(distname, c("binom", "nbinom", "geom", 
            "hyper", "pois"))) 
            discrete <- TRUE
        else discrete <- FALSE
    }
    if (!is.logical(discrete)) 
        stop("wrong argument 'discrete'.")
    if (!is.logical(keepdata) || !is.numeric(keepdata.nb) || 
        keepdata.nb < 2) 
        stop("wrong arguments 'keepdata' and 'keepdata.nb'")
    if (!is.logical(calcvcov) || length(calcvcov) != 1) 
        stop("wrong argument 'calcvcov'.")
    if (any(method == "mom")) 
        warning("the name \"mom\" for matching moments is NO MORE used and is replaced by \"mme\"")
    method <- match.arg(method, c("mle", "mme", "qme", "mge", 
        "mse"))
    if (method %in% c("mle", "mme", "mge", "mse")) 
        dpq2test <- c("d", "p")
    else dpq2test <- c("d", "p", "q")
    if (!(is.vector(data) & is.numeric(data) & length(data) > 
        1)) 
        stop("data must be a numeric vector of length greater than 1")
    checkUncensoredNAInfNan(data)
    my3dots <- list(...)
    if (length(my3dots) == 0) 
        my3dots <- NULL
    n <- length(data)
    arg_startfix <- manageparam(start.arg = start, fix.arg = fix.arg, 
        obs = data, distname = distname)
    argddistname <- names(formals(ddistname))
    hasnodefaultval <- sapply(formals(ddistname), is.name)
    arg_startfix <- checkparamlist(arg_startfix$start.arg, arg_startfix$fix.arg, 
        argddistname, hasnodefaultval)
    if (is.function(fix.arg)) 
        fix.arg.fun <- fix.arg
    else fix.arg.fun <- NULL
    resdpq <- testdpqfun(distname, dpq2test, start.arg = arg_startfix$start.arg, 
        fix.arg = arg_startfix$fix.arg, discrete = discrete)
    if (any(!resdpq$ok)) {
        for (x in resdpq[!resdpq$ok, "txt"]) warning(x)
    }
    if (method == "mme") {
        if (!is.element(distname, c("norm", "lnorm", "pois", 
            "exp", "gamma", "nbinom", "geom", "beta", "unif", 
            "logis"))) 
            if (!"order" %in% names(my3dots)) 
                stop("moment matching estimation needs an 'order' argument")
        mme <- mmedist(data, distname, start = arg_startfix$start.arg, 
            fix.arg = arg_startfix$fix.arg, checkstartfix = TRUE, 
            calcvcov = calcvcov, ...)
        varcovar <- mme$vcov
        if (!is.null(varcovar)) {
            correl <- cov2cor(varcovar)
            sd <- sqrt(diag(varcovar))
        }
        else correl <- sd <- NULL
        estimate <- mme$estimate
        loglik <- mme$loglik
        npar <- length(estimate)
        aic <- -2 * loglik + 2 * npar
        bic <- -2 * loglik + log(n) * npar
        convergence <- mme$convergence
        fix.arg <- mme$fix.arg
        weights <- mme$weights
    }
    else if (method == "mle") {
        mle <- mledist(data, distname, start = arg_startfix$start.arg, 
            fix.arg = arg_startfix$fix.arg, checkstartfix = TRUE, 
            calcvcov = calcvcov, ...)
        if (mle$convergence > 0) 
            stop("the function mle failed to estimate the parameters, \n                with the error code ", 
                mle$convergence, "\n")
        estimate <- mle$estimate
        varcovar <- mle$vcov
        if (!is.null(varcovar)) {
            correl <- cov2cor(varcovar)
            sd <- sqrt(diag(varcovar))
        }
        else correl <- sd <- NULL
        loglik <- mle$loglik
        npar <- length(estimate)
        aic <- -2 * loglik + 2 * npar
        bic <- -2 * loglik + log(n) * npar
        convergence <- mle$convergence
        fix.arg <- mle$fix.arg
        weights <- mle$weights
    }
    else if (method == "qme") {
        if (!"probs" %in% names(my3dots)) 
            stop("quantile matching estimation needs an 'probs' argument")
        qme <- qmedist(data, distname, start = arg_startfix$start.arg, 
            fix.arg = arg_startfix$fix.arg, checkstartfix = TRUE, 
            calcvcov = calcvcov, ...)
        estimate <- qme$estimate
        loglik <- qme$loglik
        npar <- length(estimate)
        aic <- -2 * loglik + 2 * npar
        bic <- -2 * loglik + log(n) * npar
        sd <- correl <- varcovar <- NULL
        convergence <- qme$convergence
        fix.arg <- qme$fix.arg
        weights <- qme$weights
    }
    else if (method == "mge") {
        if (!"gof" %in% names(my3dots)) 
            warning("maximum GOF estimation has a default 'gof' argument set to 'CvM'")
        mge <- mgedist(data, distname, start = arg_startfix$start.arg, 
            fix.arg = arg_startfix$fix.arg, checkstartfix = TRUE, 
            calcvcov = calcvcov, ...)
        estimate <- mge$estimate
        loglik <- mge$loglik
        npar <- length(estimate)
        aic <- -2 * loglik + 2 * npar
        bic <- -2 * loglik + log(n) * npar
        sd <- correl <- varcovar <- NULL
        convergence <- mge$convergence
        fix.arg <- mge$fix.arg
        weights <- NULL
    }
    else if (method == "mse") {
        mse <- msedist(data, distname, start = arg_startfix$start.arg, 
            fix.arg = arg_startfix$fix.arg, checkstartfix = TRUE, 
            calcvcov = calcvcov, ...)
        estimate <- mse$estimate
        loglik <- mse$loglik
        npar <- length(estimate)
        aic <- -2 * loglik + 2 * npar
        bic <- -2 * loglik + log(n) * npar
        sd <- correl <- varcovar <- NULL
        convergence <- mse$convergence
        fix.arg <- mse$fix.arg
        weights <- mse$weights
    }
    else {
        stop("match.arg() for 'method' did not work correctly")
    }
    if (!is.null(fix.arg)) 
        fix.arg <- as.list(fix.arg)
    if (keepdata) {
        reslist <- list(estimate = estimate, method = method, 
            sd = sd, cor = correl, vcov = varcovar, loglik = loglik, 
            aic = aic, bic = bic, n = n, data = data, distname = distname, 
            fix.arg = fix.arg, fix.arg.fun = fix.arg.fun, dots = my3dots, 
            convergence = convergence, discrete = discrete, weights = weights)
    }
    else {
        n2keep <- min(keepdata.nb, n) - 2
        imin <- which.min(data)
        imax <- which.max(data)
        subdata <- data[sample((1:n)[-c(imin, imax)], size = n2keep, 
            replace = FALSE)]
        subdata <- c(subdata, data[c(imin, imax)])
        reslist <- list(estimate = estimate, method = method, 
            sd = sd, cor = correl, vcov = varcovar, loglik = loglik, 
            aic = aic, bic = bic, n = n, data = subdata, distname = distname, 
            fix.arg = fix.arg, fix.arg.fun = fix.arg.fun, dots = my3dots, 
            convergence = convergence, discrete = discrete, weights = weights)
    }
    return(structure(reslist, class = "fitdist"))
}
