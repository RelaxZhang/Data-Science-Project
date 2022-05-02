library(MTS)
phi=matrix(c(0.2,-0.6,0.3,1.1),2,2); theta=matrix(c(-0.5,0,0,-0.5),2,2)
sigma=diag(2)
m1=VARMAsim(300,arlags=c(1),malags=c(1),phi=phi,theta=theta,sigma=sigma)
zt=m1$series
m2=VARMA(zt,p=1,q=1,include.mean=FALSE)
VARMApred(m2, h = 1, orig = 0)

X = as.matrix(fpp::usconsumption)
index_train <- 1:floor(nrow(X) * 0.8)

X_train <- X[index_train, ]
X_test <- X[-index_train, ]

obj <- nnetsauce::sklearn$linear_model$BayesianRidge()
fit_obj2 <- nnetsauce::MTS(obj = obj)

fit_obj2$fit(X_train)
preds <- fit_obj2$predict(h = nrow(X_test),level=95L, return_std = TRUE)