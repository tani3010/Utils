#include <RcppArmadillo.h>

// [[Rcpp::depends(RcppArmadillo)]]

// [[Rcpp::export]]
Rcpp::DataFrame r2norm(
    double n, double mu1, double mu2,
    double sigma1, double sigma2, double rho) {

    Rcpp::RNGScope scope;
    arma::vec tmp = Rcpp::as<arma::vec>(Rcpp::rnorm(n, 0, 1));
    arma::vec tmp2 = Rcpp::as<arma::vec>(Rcpp::rnorm(n, 0, 1));
    arma::vec x = mu1 + sigma1 * tmp;
    arma::vec y = mu2 + sigma2 * (rho * tmp + sqrt(1 - rho * rho) * tmp2);
    return Rcpp::DataFrame::create(
        Rcpp::Named("x") = Rcpp::wrap(x),
        Rcpp::Named("y") = Rcpp::wrap(y));
}

/*** R
r2norm(10, 1, 1, 0.1, 0.2, 0.3)
*/
