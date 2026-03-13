library(meta)

# 1列目を study として読み込む
dat <- read.csv("sample.csv", check.names = FALSE)

# 列名を確認
print(dat)

# 列名を解析用に変更
colnames(dat) <- c("study", "event.e", "n.e", "event.c", "n.c")

# 数値列を明示的に numeric にする
dat$event.e <- as.numeric(dat$event.e)
dat$n.e     <- as.numeric(dat$n.e)
dat$event.c <- as.numeric(dat$event.c)
dat$n.c     <- as.numeric(dat$n.c)

exp_label <- "Group A"
ctrl_label <- "Group B"

m <- metabin(
  event.e, n.e,
  event.c, n.c,
  data = dat,
  studlab = study,
  sm = "OR",
  method = "MH",
  random = TRUE,
  common = TRUE,
  incr = 0.5,
  allstudies = TRUE,
  method.random.ci = "HK"
)

summary(m)

forest(
  m,
  sortvar = NULL,
  comb.fixed = FALSE,
  comb.random = TRUE,
  leftcols = c("studlab", "event.e", "n.e", "event.c", "n.c"),
  leftlabs = c("Study", "Events", "Total", "Events", "Total"),
  lab.e = exp_label,
  lab.c = ctrl_label,
  xlab = "Odds Ratio"
)