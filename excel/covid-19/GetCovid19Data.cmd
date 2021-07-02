@Echo Off

Rem https://www3.nhk.or.jp/news/special/coronavirus/data/
    curl -L -# -O -R -S "https://www3.nhk.or.jp/n-data/opendata/coronavirus/nhk_news_covid19_prefectures_daily_data.csv"
    curl -L -# -O -R -S "https://www3.nhk.or.jp/n-data/opendata/coronavirus/nhk_news_covid19_domestic_daily_data.csv"

Rem https://www.mhlw.go.jp/stf/covid-19/open-data.html
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/severe_daily.csv"
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/cases_total.csv"
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/recovery_total.csv"
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/pcr_case_daily.csv"
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/death_total.csv"

    WScript ./ConvCovid19Data.vbs
