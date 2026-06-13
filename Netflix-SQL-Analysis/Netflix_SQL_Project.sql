-- ============================================================
--   PROJECT 03: NETFLIX CONTENT ANALYTICS DASHBOARD
--   Tool: SQL Server Compatible
--   Dataset: 03_Netflix_Content_Dataset.xlsx
-- ============================================================
--   SECTIONS:
--   1.  Database & Schema Setup
--   2.  Basic Exploration
--   3.  Content Library Analysis
--   4.  Genre Analysis
--   5.  Country & Language Analysis
--   6.  IMDb & Ratings Analysis
--   7.  Content Strategy & Trends
--   8.  Audience & Maturity Ratings
--   9.  Advanced: Window Functions
--   10. Advanced: CTEs & Subqueries
--   11. KPI Summary Dashboard Queries
-- ============================================================


-- ============================================================
-- SECTION 1: DATABASE & SCHEMA SETUP
-- ============================================================

CREATE DATABASE netflix_db;
USE netflix_db;

-- TABLE 1: Content Library
DROP TABLE IF EXISTS content_library;
CREATE TABLE content_library (
    show_id                 VARCHAR(10)     PRIMARY KEY,
    title                   VARCHAR(200)    NOT NULL,
    content_type            VARCHAR(20),
    genre                   VARCHAR(50),
    sub_genre               VARCHAR(50),
    country                 VARCHAR(50),
    language                VARCHAR(30),
    release_year            INT,
    date_added              DATE,
    content_rating          VARCHAR(10),
    duration_min            INT,
    seasons                 INT,
    director                VARCHAR(100),
    lead_actor              VARCHAR(100),
    imdb_score              DECIMAL(4,1),
    imdb_votes              INT,
    trending_score          DECIMAL(5,2),
    available_in_countries  INT,
    description_words       INT
);

-- TABLE 2: Genre Year Trends
DROP TABLE IF EXISTS genre_year_trends;
CREATE TABLE genre_year_trends (
    year                    INT,
    genre                   VARCHAR(50),
    new_titles_added        INT,
    total_on_platform       INT,
    avg_imdb_score          DECIMAL(4,2),
    total_votes             INT,
    countries_producing     INT,
    PRIMARY KEY (year, genre)
);

-- TABLE 3: Country Analysis
DROP TABLE IF EXISTS country_analysis;
CREATE TABLE country_analysis (
    country                 VARCHAR(50)     PRIMARY KEY,
    total_titles            INT,
    movies                  INT,
    tv_shows                INT,
    top_genre               VARCHAR(50),
    avg_imdb                DECIMAL(4,2),
    netflix_originals       INT,
    licensed_content        INT,
    avg_duration_min        INT,
    subscriber_est_m        DECIMAL(6,1)
);

-- TABLE 4: Ratings Distribution
DROP TABLE IF EXISTS ratings_distribution;
CREATE TABLE ratings_distribution (
    rating          VARCHAR(10)     PRIMARY KEY,
    total_titles    INT,
    movies_count    INT,
    tv_shows_count  INT,
    avg_imdb        DECIMAL(4,2),
    pct_of_library  DECIMAL(5,1)
);

-- IMPORT INSTRUCTIONS:
-- Export each Excel sheet to CSV, then use LOAD DATA INFILE


-- ============================================================
-- SECTION 2: BASIC DATA EXPLORATION
-- ============================================================

-- 2.1 Table previews
SELECT * FROM content_library;
SELECT * FROM genre_year_trends;
SELECT * FROM country_analysis;
SELECT * FROM ratings_distribution;

-- 2.2 Record counts
SELECT 'content_library'     AS tbl, COUNT(*) AS rows FROM content_library
UNION ALL
SELECT 'genre_year_trends',            COUNT(*) FROM genre_year_trends
UNION ALL
SELECT 'country_analysis',             COUNT(*) FROM country_analysis
UNION ALL
SELECT 'ratings_distribution',         COUNT(*) FROM ratings_distribution;

-- 2.3 Content type split
SELECT content_type, COUNT(*) AS titles,
       ROUND(COUNT(*)*100.0/(SELECT COUNT(*) FROM content_library),2) AS pct
FROM content_library
GROUP BY content_type;

-- 2.4 Release year range
SELECT MIN(release_year) AS earliest, MAX(release_year) AS latest,
       MIN(date_added) AS first_added, MAX(date_added) AS last_added
FROM content_library;

-- 2.5 Distinct genres and sub-genres
SELECT DISTINCT genre FROM content_library ORDER BY genre;
SELECT genre, COUNT(DISTINCT sub_genre) AS sub_genres FROM content_library GROUP BY genre;

-- 2.6 Languages available
SELECT language, COUNT(*) AS titles
FROM content_library
GROUP BY language
ORDER BY titles DESC;


-- ============================================================
-- SECTION 3: CONTENT LIBRARY ANALYSIS
-- ============================================================

-- 3.1 Movies vs TV Shows: key stats
SELECT
    content_type,
    COUNT(*)                                AS total_titles,
    ROUND(AVG(imdb_score),2)               AS avg_imdb,
    ROUND(AVG(CASE WHEN content_type='Movie' THEN duration_min END),1)  AS avg_duration_min,
    ROUND(AVG(CASE WHEN content_type='TV Show' THEN seasons END),1)     AS avg_seasons,
    ROUND(AVG(trending_score),2)           AS avg_trending_score,
    ROUND(AVG(available_in_countries),1)   AS avg_reach_countries
FROM content_library
GROUP BY content_type;

-- 3.2 Content added by year
SELECT
    YEAR(date_added)        AS year_added,
    COUNT(*)                AS titles_added,
    SUM(CASE WHEN content_type='Movie'   THEN 1 ELSE 0 END) AS movies_added,
    SUM(CASE WHEN content_type='TV Show' THEN 1 ELSE 0 END) AS shows_added,
    ROUND(AVG(imdb_score),2) AS avg_imdb_added
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY YEAR(date_added)
ORDER BY YEAR(date_added);

-- 3.3 Content added by month (seasonality)
SELECT
    FORMAT(date_added,'yyyy-MM')     AS [month],
    COUNT(*)                            AS titles_added,
    ROUND(AVG(imdb_score),2)            AS avg_imdb
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY FORMAT(date_added,'yyyy-MM')
ORDER BY FORMAT(date_added,'yyyy-MM');

-- 3.4 Time lag: release year vs date added (how old is new content?)
SELECT
    content_type,
    ROUND(AVG(YEAR(date_added) - release_year), 1) AS avg_years_lag,
    MIN(YEAR(date_added) - release_year)            AS min_lag,
    MAX(YEAR(date_added) - release_year)            AS max_lag,
    SUM(CASE WHEN YEAR(date_added) = release_year THEN 1 ELSE 0 END) AS same_year_adds
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY content_type;

-- 3.5 Multi-season TV shows (depth analysis)
SELECT
    seasons,
    COUNT(*)                            AS shows_count,
    ROUND(AVG(imdb_score),2)           AS avg_imdb,
    ROUND(AVG(trending_score),2)       AS avg_trending
FROM content_library
WHERE content_type = 'TV Show' AND seasons > 0
GROUP BY seasons
ORDER BY seasons;

-- 3.6 Long-form content vs short-form (movies)
SELECT
    CASE
        WHEN duration_min < 60  THEN 'Short (<60 min)'
        WHEN duration_min < 100 THEN 'Standard (60-99 min)'
        WHEN duration_min < 130 THEN 'Feature (100-129 min)'
        ELSE                         'Epic (130+ min)'
    END AS duration_bucket,
    COUNT(*)                            AS movies,
    ROUND(AVG(imdb_score),2)           AS avg_imdb,
    ROUND(AVG(trending_score),2)       AS avg_trending
FROM content_library
WHERE content_type = 'Movie' AND duration_min > 0
GROUP BY
    CASE
        WHEN duration_min < 60  THEN 'Short (<60 min)'
        WHEN duration_min < 100 THEN 'Standard (60-99 min)'
        WHEN duration_min < 130 THEN 'Feature (100-129 min)'
        ELSE                         'Epic (130+ min)'
    END
ORDER BY movies DESC;


-- ============================================================
-- SECTION 4: GENRE ANALYSIS
-- ============================================================

-- 4.1 Genre popularity by title count
SELECT
    genre,
    COUNT(*)                            AS total_titles,
    SUM(CASE WHEN content_type='Movie'   THEN 1 ELSE 0 END) AS movies,
    SUM(CASE WHEN content_type='TV Show' THEN 1 ELSE 0 END) AS shows,
    ROUND(AVG(imdb_score),2)           AS avg_imdb,
    ROUND(AVG(trending_score),2)       AS avg_trending,
    ROUND(AVG(available_in_countries),1) AS avg_reach
FROM content_library
GROUP BY genre
ORDER BY total_titles DESC;

-- 4.2 Sub-genre breakdown
SELECT
    genre, sub_genre,
    COUNT(*)                            AS titles,
    ROUND(AVG(imdb_score),2)           AS avg_imdb,
    ROUND(AVG(trending_score),2)       AS avg_trending
FROM content_library
GROUP BY genre, sub_genre
ORDER BY genre, titles DESC;

-- 4.3 Genre growth over years (from genre_year_trends)
SELECT
    genre,
    SUM(new_titles_added)               AS total_added,
    ROUND(AVG(avg_imdb_score),2)       AS avg_imdb_over_years,
    SUM(total_votes)                    AS total_audience_votes,
    MAX(total_on_platform)              AS current_library_size
FROM genre_year_trends
GROUP BY genre
ORDER BY total_added DESC;

-- 4.4 Fastest growing genres (2020 vs 2024)
SELECT
    genre,
    SUM(CASE WHEN year = 2020 THEN new_titles_added ELSE 0 END) AS added_2020,
    SUM(CASE WHEN year = 2024 THEN new_titles_added ELSE 0 END) AS added_2024,
    ROUND((SUM(CASE WHEN year=2024 THEN new_titles_added ELSE 0 END) -
           SUM(CASE WHEN year=2020 THEN new_titles_added ELSE 0 END)) * 100.0 /
          NULLIF(SUM(CASE WHEN year=2020 THEN new_titles_added ELSE 0 END),0), 2) AS growth_pct
FROM genre_year_trends
GROUP BY genre
ORDER BY growth_pct DESC;

-- 4.5 Genre diversity by country
SELECT TOP 15
    country,
    COUNT(DISTINCT genre)               AS genres_produced,
    COUNT(*)                            AS total_titles,
    ROUND(AVG(imdb_score),2)           AS avg_imdb
FROM content_library
GROUP BY country
ORDER BY genres_produced DESC;


-- ============================================================
-- SECTION 5: COUNTRY & LANGUAGE ANALYSIS
-- ============================================================

-- 5.1 Top content-producing countries
SELECT TOP 15
    country,
    total_titles,
    movies,
    tv_shows,
    top_genre,
    avg_imdb,
    netflix_originals,
    ROUND(netflix_originals*100.0/total_titles,2) AS originals_pct,
    subscriber_est_m
FROM country_analysis
ORDER BY total_titles DESC;

-- 5.2 Language distribution
SELECT
    language,
    COUNT(*)                            AS titles,
    ROUND(AVG(imdb_score),2)           AS avg_imdb,
    ROUND(AVG(trending_score),2)       AS avg_trending,
    ROUND(COUNT(*)*100.0/(SELECT COUNT(*) FROM content_library),2) AS library_pct
FROM content_library
GROUP BY language
ORDER BY titles DESC;

-- 5.3 Country IMDb comparison
SELECT
    ca.country,
    ca.total_titles,
    ca.avg_imdb,
    ca.subscriber_est_m,
    ROUND(ca.avg_imdb - (SELECT AVG(imdb_score) FROM content_library),2) AS vs_global_avg
FROM country_analysis ca
ORDER BY ca.avg_imdb DESC;

-- 5.4 Netflix Originals by country
SELECT TOP 15
    country,
    netflix_originals,
    total_titles,
    ROUND(netflix_originals*100.0/total_titles,2) AS originals_pct
FROM country_analysis
ORDER BY netflix_originals DESC;

-- 5.5 Global reach analysis (available_in_countries)
SELECT
    content_type, genre,
    ROUND(AVG(available_in_countries),1)    AS avg_reach,
    MAX(available_in_countries)             AS max_reach,
    MIN(available_in_countries)             AS min_reach,
    COUNT(*)                                AS titles
FROM content_library
GROUP BY content_type, genre
ORDER BY avg_reach DESC;


-- ============================================================
-- SECTION 6: IMDB & RATINGS ANALYSIS
-- ============================================================

-- 6.1 IMDb score distribution
SELECT
    CASE
        WHEN imdb_score >= 8.0 THEN 'Excellent (8.0+)'
        WHEN imdb_score >= 7.0 THEN 'Good (7.0-7.9)'
        WHEN imdb_score >= 6.0 THEN 'Average (6.0-6.9)'
        WHEN imdb_score >= 5.0 THEN 'Below Average (5.0-5.9)'
        ELSE                        'Poor (<5.0)'
    END AS rating_tier,
    COUNT(*)                            AS titles,
    ROUND(AVG(trending_score),2)       AS avg_trending,
    ROUND(AVG(available_in_countries),1) AS avg_reach
FROM content_library
GROUP BY
    CASE
        WHEN imdb_score >= 8.0 THEN 'Excellent (8.0+)'
        WHEN imdb_score >= 7.0 THEN 'Good (7.0-7.9)'
        WHEN imdb_score >= 6.0 THEN 'Average (6.0-6.9)'
        WHEN imdb_score >= 5.0 THEN 'Below Average (5.0-5.9)'
        ELSE                        'Poor (<5.0)'
    END
ORDER BY titles DESC;

-- 6.2 Highly rated content by genre
SELECT
    genre,
    COUNT(CASE WHEN imdb_score >= 8.0 THEN 1 END) AS excellent_titles,
    COUNT(*) AS total_titles,
    ROUND(COUNT(CASE WHEN imdb_score>=8.0 THEN 1 END)*100.0/COUNT(*),2) AS excellent_pct,
    ROUND(AVG(imdb_score),2) AS avg_imdb
FROM content_library
GROUP BY genre
ORDER BY excellent_pct DESC;

-- 6.3 Most voted (popular) content
SELECT TOP 20
    title, content_type, genre, country, release_year,
    imdb_score, imdb_votes, trending_score
FROM content_library
ORDER BY imdb_votes DESC;

-- 6.4 Hidden gems: High IMDb but low votes (underrated)
SELECT TOP 20
    title, content_type, genre, country, release_year,
    imdb_score, imdb_votes, trending_score
FROM content_library
WHERE imdb_score >= 7.5 AND imdb_votes < 10000
ORDER BY imdb_score DESC;

-- 6.5 Content rating (maturity) vs IMDb quality
SELECT
    content_rating,
    COUNT(*)                            AS titles,
    ROUND(AVG(imdb_score),2)           AS avg_imdb,
    ROUND(AVG(trending_score),2)       AS avg_trending,
    ROUND(AVG(available_in_countries),1) AS avg_global_reach
FROM content_library
GROUP BY content_rating
ORDER BY avg_imdb DESC;

-- 6.6 Correlation: IMDb vs Trending Score (buckets)
SELECT
    ROUND(imdb_score) AS imdb_rounded,
    COUNT(*)                            AS titles,
    ROUND(AVG(trending_score),2)       AS avg_trending
FROM content_library
GROUP BY ROUND(imdb_score,0)
ORDER BY ROUND(imdb_score,0) DESC;


-- ============================================================
-- SECTION 7: CONTENT STRATEGY & TRENDS
-- ============================================================

-- 7.1 What genres are trending (high trending score)?
SELECT
    genre,
    ROUND(AVG(trending_score),2)       AS avg_trending,
    COUNT(*)                            AS titles,
    ROUND(AVG(imdb_score),2)           AS avg_imdb
FROM content_library
GROUP BY genre
ORDER BY avg_trending DESC;

-- 7.2 Recent additions (last 2 years) by genre
SELECT
    genre,
    COUNT(*) AS recent_additions
FROM content_library
WHERE date_added >= DATEADD(YEAR,-2,GETDATE())
GROUP BY genre
ORDER BY recent_additions DESC;

-- 7.3 Content strategy by year: original vs catalog
SELECT
    YEAR(date_added) AS year_added,
    ROUND(AVG(YEAR(date_added) - release_year),1) AS avg_content_age,
    COUNT(CASE WHEN YEAR(date_added) = release_year THEN 1 END) AS fresh_content,
    COUNT(CASE WHEN YEAR(date_added) - release_year > 5 THEN 1 END) AS catalog_content,
    COUNT(*) AS total
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY YEAR(date_added)
ORDER BY YEAR(date_added);

-- 7.4 Genre IMDb trend over years
SELECT
    year,
    genre,
    avg_imdb_score,
    LAG(avg_imdb_score) OVER (PARTITION BY genre ORDER BY [year]) AS prev_year_imdb,
    ROUND(avg_imdb_score - LAG(avg_imdb_score) OVER (PARTITION BY genre ORDER BY [year]),2) AS imdb_change
FROM genre_year_trends
ORDER BY genre, year;


-- ============================================================
-- SECTION 8: AUDIENCE & MATURITY RATINGS
-- ============================================================

-- 8.1 Rating distribution summary
SELECT
    rating,
    total_titles,
    movies_count,
    tv_shows_count,
    avg_imdb,
    pct_of_library
FROM ratings_distribution
ORDER BY total_titles DESC;

-- 8.2 Content rating by genre
SELECT
    genre, content_rating,
    COUNT(*) AS titles,
    ROUND(COUNT(*)*100.0/SUM(COUNT(*)) OVER (PARTITION BY genre),2) AS pct_within_genre
FROM content_library
GROUP BY genre, content_rating
ORDER BY genre, titles DESC;

-- 8.3 Family-friendly content analysis
SELECT
    country,
    SUM(CASE WHEN content_rating IN ('G','PG','TV-PG','TV-Y7') THEN 1 ELSE 0 END) AS family_friendly,
    COUNT(*) AS total,
    ROUND(SUM(CASE WHEN content_rating IN ('G','PG','TV-PG','TV-Y7') THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS family_pct
FROM content_library
GROUP BY country
HAVING total >= 20
ORDER BY family_pct DESC;

-- 8.4 Adult content by country
SELECT TOP 15
    country,
    SUM(CASE WHEN content_rating IN ('R','TV-MA') THEN 1 ELSE 0 END) AS adult_titles,
    COUNT(*) AS total,
    ROUND(SUM(CASE WHEN content_rating IN ('R','TV-MA') THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS adult_pct
FROM content_library
GROUP BY country
HAVING total >= 20
ORDER BY adult_pct DESC;


-- ============================================================
-- SECTION 9: WINDOW FUNCTIONS
-- ============================================================

-- 9.1 Rank content within each genre by IMDb score
SELECT TOP 50
    title, genre, imdb_score, country,
    RANK() OVER (PARTITION BY genre ORDER BY imdb_score DESC)        AS genre_rank,
    DENSE_RANK() OVER (PARTITION BY genre ORDER BY imdb_score DESC)  AS genre_dense_rank
FROM content_library
ORDER BY genre, genre_rank;

-- 9.2 Cumulative titles added to Netflix per genre
SELECT
    YEAR(date_added) AS year_added, genre,
    COUNT(*) AS added_this_year,
    SUM(COUNT(*)) OVER (PARTITION BY genre ORDER BY YEAR(date_added)
                        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS cumulative_titles
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY YEAR(date_added), genre
ORDER BY genre, YEAR(date_added);

-- 9.3 IMDb score percentile within content type
SELECT
    title, content_type, genre, imdb_score,
    ROUND(PERCENT_RANK() OVER (PARTITION BY content_type ORDER BY imdb_score) * 100, 2) AS imdb_percentile,
    NTILE(10) OVER (PARTITION BY content_type ORDER BY imdb_score) AS imdb_decile
FROM content_library
ORDER BY content_type, imdb_score DESC;

-- 9.4 Month-over-month content additions
SELECT
    FORMAT(date_added,'yyyy-MM') AS [month],
    COUNT(*) AS added,
    LAG(COUNT(*)) OVER (ORDER BY FORMAT(date_added,'yyyy-MM')) AS prev_month,
    COUNT(*) - LAG(COUNT(*)) OVER (ORDER BY FORMAT(date_added,'yyyy-MM')) AS mom_change
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY FORMAT(date_added,'yyyy-MM')
ORDER BY FORMAT(date_added,'yyyy-MM');

-- 9.5 Top 3 titles per country by IMDb score
SELECT * FROM (
    SELECT
        title, country, genre, imdb_score, content_type,
        ROW_NUMBER() OVER (PARTITION BY country ORDER BY imdb_score DESC) AS country_rank
    FROM content_library
) ranked
WHERE country_rank <= 3
ORDER BY country, country_rank;

-- 9.6 Running average IMDb score per genre over time
SELECT
    YEAR(date_added) AS year_added, genre,
    ROUND(AVG(imdb_score),2) AS year_avg_imdb,
    ROUND(AVG(AVG(imdb_score)) OVER (PARTITION BY genre ORDER BY YEAR(date_added)
          ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW),2) AS cumulative_avg_imdb
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY YEAR(date_added), genre
ORDER BY genre, YEAR(date_added);

-- 9.7 Genre share of content library per year (PARTITION BY)
SELECT
    YEAR(date_added) AS year_added,
    genre,
    COUNT(*) AS genre_count,
    ROUND(COUNT(*)*100.0/SUM(COUNT(*)) OVER (PARTITION BY YEAR(date_added)),2) AS genre_share_pct
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY YEAR(date_added), genre
ORDER BY YEAR(date_added), genre_count DESC;


-- ============================================================
-- SECTION 10: CTEs & SUBQUERIES
-- ============================================================

;
-- 10.1 CTE: Content quality tiers
WITH quality_tiers AS (
    SELECT *,
        CASE
            WHEN imdb_score >= 8.0 AND imdb_votes > 50000 THEN 'Blockbuster'
            WHEN imdb_score >= 7.0 AND imdb_votes > 20000 THEN 'Mainstream Hit'
            WHEN imdb_score >= 7.0                         THEN 'Critic Favorite'
            WHEN imdb_score >= 5.5                         THEN 'Average'
            ELSE                                                'Low Quality'
        END AS quality_tier
    FROM content_library
)
SELECT
    quality_tier,
    COUNT(*)                                AS titles,
    ROUND(AVG(trending_score),2)           AS avg_trending,
    ROUND(AVG(available_in_countries),1)   AS avg_reach,
    STRING_AGG(genre, ', ') AS genres
FROM quality_tiers
GROUP BY
    CASE
        WHEN imdb_score >= 8.0 AND imdb_votes > 50000 THEN 'Blockbuster'
        WHEN imdb_score >= 7.0 AND imdb_votes > 20000 THEN 'Mainstream Hit'
        WHEN imdb_score >= 7.0                         THEN 'Critic Favorite'
        WHEN imdb_score >= 5.5                         THEN 'Average'
        ELSE                                                'Low Quality'
    END
ORDER BY titles DESC;

;
-- 10.2 CTE: Genre-country content matrix
WITH genre_country AS (
    SELECT genre, country, COUNT(*) AS titles
    FROM content_library
    GROUP BY genre, country
),
ranked AS (
    SELECT *,
        RANK() OVER (PARTITION BY genre ORDER BY titles DESC) AS country_rank
    FROM genre_country
)
SELECT genre, country, titles, country_rank
FROM ranked
WHERE country_rank <= 3
ORDER BY genre, country_rank;

;
-- 10.3 CTE: Netflix content investment signals
WITH content_signals AS (
    SELECT
        genre,
        ROUND(AVG(imdb_score),2)           AS avg_imdb,
        ROUND(AVG(trending_score),2)       AS avg_trending,
        ROUND(AVG(available_in_countries),1) AS avg_reach,
        COUNT(*)                            AS total_titles,
        SUM(CASE WHEN YEAR(date_added)>=2022 THEN 1 ELSE 0 END) AS recent_titles
    FROM content_library
    GROUP BY genre
),
scored AS (
    SELECT *,
        ROUND((avg_imdb * 10) + (avg_trending * 5) + (avg_reach / 10) + (recent_titles * 2), 2) AS invest_score
    FROM content_signals
)
SELECT genre, avg_imdb, avg_trending, avg_reach, total_titles, recent_titles, invest_score,
       RANK() OVER (ORDER BY invest_score DESC) AS invest_priority_rank
FROM scored
ORDER BY invest_score DESC;

;
-- 10.4 CTE: Year-over-year genre growth
WITH yearly_counts AS (
    SELECT YEAR(date_added) AS yr, genre, COUNT(*) AS titles
    FROM content_library WHERE date_added IS NOT NULL
    GROUP BY YEAR(date_added), genre
),
with_lag AS (
    SELECT *,
        LAG(titles) OVER (PARTITION BY genre ORDER BY yr) AS prev_yr_titles
    FROM yearly_counts
)
SELECT yr, genre, titles, prev_yr_titles,
       ROUND((titles - prev_yr_titles)*100.0/NULLIF(prev_yr_titles,0),2) AS yoy_growth_pct
FROM with_lag
WHERE prev_yr_titles IS NOT NULL
ORDER BY genre, yr;

-- 10.5 Subquery: Countries producing above-average quality content
SELECT
    country,
    ROUND(AVG(imdb_score),2) AS avg_imdb,
    COUNT(*) AS titles
FROM content_library
GROUP BY country
HAVING AVG(imdb_score) > (SELECT AVG(imdb_score) FROM content_library)
   AND COUNT(*) >= 10
ORDER BY avg_imdb DESC;

;
-- 10.6 CTE: Multi-factor content recommendation engine
WITH scored_content AS (
    SELECT
        show_id, title, content_type, genre, country, language,
        release_year, imdb_score, imdb_votes, trending_score, available_in_countries,
        ROUND(
            (imdb_score * 10) +
            (LOG(imdb_votes + 1) * 5) +
            (trending_score * 8) +
            (available_in_countries / 10),
        2) AS recommendation_score
    FROM content_library
)
SELECT TOP 30
    title, content_type, genre, country, language, release_year,
    imdb_score, trending_score, recommendation_score,
    RANK() OVER (PARTITION BY genre ORDER BY recommendation_score DESC) AS genre_rec_rank
FROM scored_content
ORDER BY recommendation_score DESC;


-- ============================================================
-- SECTION 11: KPI SUMMARY DASHBOARD QUERIES
-- ============================================================

-- 11.1 Platform content overview
SELECT
    COUNT(*)                                AS total_titles,
    SUM(CASE WHEN content_type='Movie'   THEN 1 ELSE 0 END) AS total_movies,
    SUM(CASE WHEN content_type='TV Show' THEN 1 ELSE 0 END) AS total_tv_shows,
    COUNT(DISTINCT genre)                   AS genres,
    COUNT(DISTINCT country)                 AS countries,
    COUNT(DISTINCT language)                AS languages,
    ROUND(AVG(imdb_score),2)               AS avg_imdb_score,
    ROUND(AVG(trending_score),2)           AS avg_trending_score,
    ROUND(AVG(available_in_countries),1)   AS avg_global_reach
FROM content_library;

-- 11.2 Genre leaderboard
SELECT
    genre,
    COUNT(*)                                AS titles,
    ROUND(AVG(imdb_score),2)               AS avg_imdb,
    ROUND(AVG(trending_score),2)           AS avg_trending,
    SUM(imdb_votes)                         AS total_votes,
    RANK() OVER (ORDER BY COUNT(*) DESC)    AS title_rank,
    RANK() OVER (ORDER BY AVG(imdb_score) DESC) AS quality_rank
FROM content_library
GROUP BY genre
ORDER BY title_rank;

-- 11.3 Country content scorecard
SELECT
    ca.country,
    ca.total_titles,
    ca.avg_imdb,
    ca.netflix_originals,
    ROUND(ca.netflix_originals*100.0/ca.total_titles,2) AS originals_pct,
    ca.subscriber_est_m,
    RANK() OVER (ORDER BY ca.total_titles DESC) AS content_volume_rank,
    RANK() OVER (ORDER BY ca.avg_imdb DESC)     AS quality_rank
FROM country_analysis ca
ORDER BY content_volume_rank;

-- 11.4 Annual content health report
SELECT
    YEAR(date_added)                        AS [year],
    COUNT(*)                                AS titles_added,
    ROUND(AVG(imdb_score),2)               AS avg_imdb,
    ROUND(AVG(trending_score),2)           AS avg_trending,
    COUNT(DISTINCT genre)                   AS genre_diversity,
    COUNT(DISTINCT country)                 AS country_diversity,
    ROUND(AVG(available_in_countries),1)   AS avg_global_reach
FROM content_library
WHERE date_added IS NOT NULL
GROUP BY YEAR(date_added)
ORDER BY YEAR(date_added);

-- ============================================================
-- END OF PROJECT 03: NETFLIX CONTENT ANALYTICS
-- ============================================================
