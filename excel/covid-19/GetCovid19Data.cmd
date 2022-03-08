@Echo Off

    if not exist .\data\. mkdir .\data
    if not exist .\conv\. mkdir .\conv

    pushd .\data

Rem https://github.com/office-itou/Files/tree/master/excel/covid-19
    curl -L -# -R -S -o "�l��(�l�����v2019).csv" "https://raw.githubusercontent.com/office-itou/Files/master/excel/covid-19/Population(2019).csv"
    curl -L -# -R -S -o "�l��(��������2020).csv" "https://raw.githubusercontent.com/office-itou/Files/master/excel/covid-19/Population(2020).csv"

Rem https://www3.nhk.or.jp/news/special/coronavirus/data/
    curl -L -# -O -R -S "https://www3.nhk.or.jp/n-data/opendata/coronavirus/nhk_news_covid19_prefectures_daily_data.csv"
    curl -L -# -O -R -S "https://www3.nhk.or.jp/n-data/opendata/coronavirus/nhk_news_covid19_domestic_daily_data.csv"

Rem https://www.mhlw.go.jp/stf/covid-19/open-data.html
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/severe_daily.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/cases_total.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/recovery_total.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/pcr_case_daily.csv"
Rem curl -L -# -O -R -S "https://www.mhlw.go.jp/content/death_total.csv"

Rem PCR�������{�l��
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/pcr_case_daily.csv"

Rem PCR�����̎��{����
    curl -L -# -O -R -S "https://www.mhlw.go.jp/content/pcr_tested_daily.csv"

Rem �V�K�z���Ґ��̐��ځi���ʁj
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_daily.csv"

Rem �l��10���l������V�K�z���Ґ�
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_per_100_thousand_population_daily.csv"

Rem ���ʁE�N��ʐV�K�z���Ґ��i�T�ʁj
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_detail_weekly.csv"

Rem �z���Ґ��i�ݐρj
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/confirmed_cases_cumulative_daily.csv"

Rem ���ʁE�N��ʗz���Ґ��i�ݐρj
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/confirmed_cases_detail_cumulative_weekly.csv"

Rem �d�ǎҐ��̐���
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/severe_cases_daily.csv"

Rem ���ʁE�N��ʏd�ǎҐ�
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/severe_cases_detail_weekly.csv"

Rem ���S�Ґ��i�ݐρj
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/deaths_cumulative_daily.csv"

Rem ���ʁE�N��ʎ��S�Ґ��i�ݐρj
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/deaths_detail_cumulative_weekly.csv"

Rem ���@���Ó���v����ғ�����
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/requiring_inpatient_care_etc_daily.csv"

Rem �W�c������������
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/cluster_events_weekly.csv"

Rem HER-SYS�f�[�^�Ɋ�Â��V�K�z���Ґ�
    curl -L -# -O -R -S "https://covid19.mhlw.go.jp/public/opendata/newly_confirmed_cases_difference.csv"

    popd

    CScript ./ConvCovid19Data.vbs

    Pause.
