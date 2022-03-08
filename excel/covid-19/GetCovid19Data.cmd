@Echo Off

    if not exist .\data\. mkdir .\data
    if not exist .\conv\. mkdir .\conv

    pushd .\data

Rem https://www3.nhk.or.jp/news/special/coronavirus/data/
    curl -L -# -O -R -S "https://www3.nhk.or.jp/n-data/opendata/coronavirus/nhk_news_covid19_prefectures_daily_data.csv"
    curl -L -# -O -R -S "https://www3.nhk.or.jp/n-data/opendata/coronavirus/nhk_news_covid19_domestic_daily_data.csv"

Rem https://www.mhlw.go.jp/stf/covid-19/open-data.html
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/severe_daily.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/cases_total.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/recovery_total.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/pcr_case_daily.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/death_total.csv"

Rem PCR検査実施人数
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/pcr_case_daily.csv"

Rem PCR検査の実施件数
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/pcr_tested_daily.csv"

Rem 新規陽性者数の推移（日別）
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_daily.csv"

Rem 人口10万人当たり新規陽性者数
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_per_100_thousand_population_daily.csv"

Rem 性別・年代別新規陽性者数（週別）
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_detail_weekly.csv"

Rem 陽性者数（累積）
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/confirmed_cases_cumulative_daily.csv"

Rem 性別・年代別陽性者数（累積）
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/confirmed_cases_detail_cumulative_weekly.csv"

Rem 重症者数の推移
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/severe_cases_daily.csv"

Rem 性別・年代別重症者数
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/severe_cases_detail_weekly.csv"

Rem 死亡者数（累積）
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/deaths_cumulative_daily.csv"

Rem 性別・年代別死亡者数（累積）
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/deaths_detail_cumulative_weekly.csv"

Rem 入院治療等を要する者等推移
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/requiring_inpatient_care_etc_daily.csv"

Rem 集団感染等発生状況
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/cluster_events_weekly.csv"

Rem HER-SYSデータに基づく新規陽性者数
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_difference.csv"

    popd

    CScript ./ConvCovid19Data.vbs

    Pause.
