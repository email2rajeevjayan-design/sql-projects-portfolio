-- ============================================================
--   PROJECT 01: HEALTHCARE OPERATIONS ANALYTICS
--   Tool: MySQL / PostgreSQL Compatible
--   Dataset: 01_Healthcare_Operations_Dataset.xlsx
-- ============================================================
--   SECTIONS:
--   1.  Database & Schema Setup
--   2.  Basic Exploration
--   3.  Patient Admissions Analysis
--   4.  Departmental Performance
--   5.  Financial & Billing Analysis
--   6.  Doctor & Staff Analysis
--   7.  Readmission & Risk Analysis
--   8.  Advanced: Window Functions
--   9.  Advanced: CTEs & Subqueries
--   10. KPI Summary Dashboard Queries
-- ============================================================


-- ============================================================
-- SECTION 1: DATABASE & SCHEMA SETUP
-- ============================================================

CREATE DATABASE  healthcare_db;
USE healthcare_db;

-- TABLE 1: Patient Admissions
DROP TABLE IF EXISTS patient_admissions;
CREATE TABLE patient_admissions (
    admission_id        VARCHAR(10)     PRIMARY KEY,
    patient_id          VARCHAR(10)     NOT NULL,
    patient_name        VARCHAR(100)    NOT NULL,
    age                 INT,
    gender              VARCHAR(10),
    blood_type          VARCHAR(5),
    admission_date      DATE,
    discharge_date      DATE,
    los_days            INT,
    department          VARCHAR(50),
    ward                VARCHAR(20),
    admission_type      VARCHAR(20),
    diagnosis           VARCHAR(100),
    procedure_done      VARCHAR(100),
    doctor_id           VARCHAR(10),
    doctor_name         VARCHAR(100),
    insurance_type      VARCHAR(20),
    billing_amount      DECIMAL(12,2),
    status              VARCHAR(20),
    readmission_flag    VARCHAR(5)
);

-- TABLE 2: Department KPIs
DROP TABLE IF EXISTS department_kpis;
CREATE TABLE department_kpis (
    month                   VARCHAR(10),
    department              VARCHAR(50),
    total_admissions        INT,
    avg_los_days            DECIMAL(5,1),
    bed_occupancy_rate_pct  DECIMAL(5,1),
    patient_satisfaction    DECIMAL(4,1),
    revenue_inr             DECIMAL(15,2),
    operating_cost_inr      DECIMAL(15,2),
    staff_count             INT,
    readmission_rate_pct    DECIMAL(5,1),
    mortality_rate_pct      DECIMAL(5,2),
    emergency_cases         INT,
    PRIMARY KEY (month, department)
);

-- TABLE 3: Staff Schedule
DROP TABLE IF EXISTS staff_schedule;
CREATE TABLE staff_schedule (
    staff_id            VARCHAR(10)     PRIMARY KEY,
    staff_name          VARCHAR(100),
    role                VARCHAR(30),
    department          VARCHAR(50),
    shift               VARCHAR(20),
    days_worked         INT,
    overtime_hours      DECIMAL(5,1),
    monthly_salary_inr  DECIMAL(10,2)
);



-- ============================================================
-- SECTION 2: BASIC DATA EXPLORATION
-- ============================================================

-- 2.1 Preview all tables
SELECT * FROM patient_admissions
SELECT * FROM department_kpis
SELECT * FROM staff_schedule

-- 2.2 Row counts
SELECT 'patient_admissions' AS table_name, COUNT(*) AS total_rows FROM patient_admissions
UNION ALL
SELECT 'department_kpis',                  COUNT(*) FROM department_kpis
UNION ALL
SELECT 'staff_schedule',                   COUNT(*) FROM staff_schedule;

-- 2.3 Date range of admissions
SELECT
    MIN(admission_date) AS earliest_admission,
    MAX(admission_date) AS latest_admission,
    DATEDIFF(DAY, MIN(admission_date), MAX(admission_date)) AS span_days
FROM patient_admissions;



-- 2.4 Distinct values check
SELECT DISTINCT department   FROM patient_admissions ORDER BY department;
SELECT DISTINCT admission_type FROM patient_admissions;
SELECT DISTINCT insurance_type FROM patient_admissions;
SELECT DISTINCT status         FROM patient_admissions;


-- ============================================================
-- SECTION 3: PATIENT ADMISSIONS ANALYSIS
-- ============================================================

-- 3.1 Total admissions by department
SELECT
    department,
    COUNT(*)                        AS total_admissions,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM patient_admissions), 2) AS pct_share
FROM patient_admissions
GROUP BY department
ORDER BY total_admissions DESC;

-- 3.2 Admissions by admission type
SELECT
    admission_type,
    COUNT(*)                            AS total,
    ROUND(AVG(los_days), 2)             AS avg_los,
    ROUND(AVG(billing_amount), 2)       AS avg_billing
FROM patient_admissions
GROUP BY admission_type
ORDER BY total DESC;

-- 3.3 Monthly admissions trend
SELECT
    FORMAT(admission_date, 'yyyy-MM') AS [month],
    COUNT(*) AS admissions,
    SUM(billing_amount) AS total_revenue
FROM patient_admissions
GROUP BY FORMAT(admission_date, 'yyyy-MM')
ORDER BY [month];

-- 3.4 Age group distribution
SELECT
    CASE
        WHEN age < 18 THEN '0-17 (Pediatric)'
        WHEN age < 35 THEN '18-34 (Young Adult)'
        WHEN age < 55 THEN '35-54 (Middle Age)'
        WHEN age < 70 THEN '55-69 (Senior)'
        ELSE '70+ (Elderly)'
    END AS age_group,
    COUNT(*) AS patient_count,
    ROUND(AVG(los_days), 2) AS avg_los,
    ROUND(AVG(billing_amount), 2) AS avg_billing
FROM patient_admissions
GROUP BY
    CASE
        WHEN age < 18 THEN '0-17 (Pediatric)'
        WHEN age < 35 THEN '18-34 (Young Adult)'
        WHEN age < 55 THEN '35-54 (Middle Age)'
        WHEN age < 70 THEN '55-69 (Senior)'
        ELSE '70+ (Elderly)'
    END
ORDER BY patient_count DESC;

-- 3.5 Gender-based analysis
SELECT
    gender,
    COUNT(*)                            AS total_patients,
    ROUND(AVG(age), 1)                  AS avg_age,
    ROUND(AVG(los_days), 2)             AS avg_los,
    ROUND(AVG(billing_amount), 2)       AS avg_billing,
    SUM(CASE WHEN readmission_flag = 'Yes' THEN 1 ELSE 0 END) AS readmissions
FROM patient_admissions
GROUP BY gender;

-- 3.6 Top 10 diagnoses
SELECT TOP 10
    diagnosis,
    COUNT(*) AS cases,
    ROUND(AVG(los_days), 2) AS avg_los,
    ROUND(AVG(billing_amount), 2) AS avg_billing
FROM patient_admissions
GROUP BY diagnosis
ORDER BY cases DESC;

-- 3.7 Blood type distribution
SELECT
    blood_type,
    COUNT(*) AS patient_count
FROM patient_admissions
GROUP BY blood_type
ORDER BY patient_count DESC;

-- 3.8 Insurance type vs billing
SELECT
    insurance_type,
    COUNT(*)                            AS patients,
    ROUND(AVG(billing_amount), 2)       AS avg_billing,
    ROUND(SUM(billing_amount), 2)       AS total_billing,
    ROUND(MIN(billing_amount), 2)       AS min_billing,
    ROUND(MAX(billing_amount), 2)       AS max_billing
FROM patient_admissions
GROUP BY insurance_type
ORDER BY total_billing DESC;

-- 3.9 Length of stay analysis
SELECT
    department,
    ROUND(AVG(los_days), 2)             AS avg_los,
    MIN(los_days)                       AS min_los,
    MAX(los_days)                       AS max_los,
    COUNT(CASE WHEN los_days > 14 THEN 1 END) AS long_stay_count
FROM patient_admissions
GROUP BY department
ORDER BY avg_los DESC;

-- 3.10 Active vs Discharged patients by department
SELECT
    department,
    SUM(CASE WHEN status = 'Active'    THEN 1 ELSE 0 END) AS active,
    SUM(CASE WHEN status = 'Discharged' THEN 1 ELSE 0 END) AS discharged,
    COUNT(*) AS total
FROM patient_admissions
GROUP BY department
ORDER BY total DESC;


-- ============================================================
-- SECTION 4: DEPARTMENTAL PERFORMANCE
-- ============================================================

-- 4.1 Average KPIs per department (all time)
SELECT
    department,
    ROUND(AVG(total_admissions), 1)       AS avg_monthly_admissions,
    ROUND(AVG(avg_los_days), 2)           AS avg_los,
    ROUND(AVG(bed_occupancy_rate_pct), 2) AS avg_occupancy_pct,
    ROUND(AVG(patient_satisfaction), 2)   AS avg_satisfaction,
    ROUND(AVG(readmission_rate_pct), 2)   AS avg_readmission_pct,
    ROUND(AVG(mortality_rate_pct), 3)     AS avg_mortality_pct
FROM department_kpis
GROUP BY department
ORDER BY avg_monthly_admissions DESC;

-- 4.2 Monthly revenue vs cost by department
SELECT
    month,
    department,
    revenue_inr,
    operating_cost_inr,
    ROUND(revenue_inr - operating_cost_inr, 2) AS net_profit,
    ROUND((revenue_inr - operating_cost_inr) / revenue_inr * 100, 2) AS profit_margin_pct
FROM department_kpis
ORDER BY month, department;

-- 4.3 Top profitable departments
SELECT
    department,
    ROUND(SUM(revenue_inr), 2)                               AS total_revenue,
    ROUND(SUM(operating_cost_inr), 2)                        AS total_cost,
    ROUND(SUM(revenue_inr - operating_cost_inr), 2)          AS total_profit,
    ROUND(AVG((revenue_inr - operating_cost_inr) / revenue_inr * 100), 2) AS avg_margin_pct
FROM department_kpis
GROUP BY department
ORDER BY total_profit DESC;

-- 4.4 Bed occupancy comparison (2023 vs 2024)
SELECT
    department,
    ROUND(AVG(CASE WHEN month LIKE '2023%' THEN bed_occupancy_rate_pct END), 2) AS avg_occupancy_2023,
    ROUND(AVG(CASE WHEN month LIKE '2024%' THEN bed_occupancy_rate_pct END), 2) AS avg_occupancy_2024,
    ROUND(AVG(CASE WHEN month LIKE '2024%' THEN bed_occupancy_rate_pct END) -
          AVG(CASE WHEN month LIKE '2023%' THEN bed_occupancy_rate_pct END), 2) AS yoy_change
FROM department_kpis
GROUP BY department
ORDER BY yoy_change DESC;

-- 4.5 Emergency cases trend by quarter
SELECT
    department,
    CONCAT(
        LEFT([month],4),
        '-Q',
        CEILING(CAST(SUBSTRING([month],6,2) AS INT) / 3.0)
    ) AS quarter,
    SUM(emergency_cases) AS total_emergencies,
    SUM(total_admissions) AS total_admissions,
    ROUND(SUM(emergency_cases) * 100.0 / SUM(total_admissions), 2) AS emergency_pct
FROM department_kpis
GROUP BY
    department,
    CONCAT(
        LEFT([month],4),
        '-Q',
        CEILING(CAST(SUBSTRING([month],6,2) AS INT) / 3.0)
    )
ORDER BY department, quarter;

-- 4.6 Patient satisfaction ranking
SELECT
    department,
    ROUND(AVG(patient_satisfaction), 2) AS avg_satisfaction,
    RANK() OVER (ORDER BY AVG(patient_satisfaction) DESC) AS satisfaction_rank
FROM department_kpis
GROUP BY department;

-- 4.7 Staff productivity: admissions per staff
SELECT
    department,
    ROUND(AVG(total_admissions), 1)          AS avg_admissions,
    ROUND(AVG(staff_count), 1)               AS avg_staff,
    ROUND(AVG(total_admissions) / AVG(staff_count), 2) AS admissions_per_staff
FROM department_kpis
GROUP BY department
ORDER BY admissions_per_staff DESC;


-- ============================================================
-- SECTION 5: FINANCIAL & BILLING ANALYSIS
-- ============================================================

-- 5.1 Total revenue by year
SELECT
    YEAR(admission_date) AS [year],
    COUNT(*) AS total_admissions,
    ROUND(SUM(billing_amount), 2) AS total_revenue,
    ROUND(AVG(billing_amount), 2) AS avg_billing
FROM patient_admissions
GROUP BY YEAR(admission_date)
ORDER BY [year];

-- 5.2 High-value patients (top 10% billing)
SELECT TOP 20
    patient_id,
    patient_name,
    department,
    diagnosis,
    billing_amount,
    los_days,
    insurance_type
FROM patient_admissions
WHERE billing_amount >= (
    SELECT DISTINCT
        PERCENTILE_CONT(0.90) WITHIN GROUP (ORDER BY billing_amount)
        OVER ()
    FROM patient_admissions
)
ORDER BY billing_amount DESC;

-- MySQL alternative for top 10% billing:
SELECT TOP 100
    patient_id,
    patient_name,
    department,
    diagnosis,
    billing_amount,
    los_days,
    insurance_type
FROM patient_admissions
ORDER BY billing_amount DESC;

-- 5.3 Revenue by diagnosis category
SELECT TOP 15
    diagnosis,
    COUNT(*) AS cases,
    ROUND(SUM(billing_amount), 2) AS total_revenue,
    ROUND(AVG(billing_amount), 2) AS avg_billing,
    ROUND(MAX(billing_amount), 2) AS max_billing
FROM patient_admissions
GROUP BY diagnosis
ORDER BY total_revenue DESC;

-- 5.4 Department-wise revenue per LOS day (efficiency metric)
SELECT
    department,
    ROUND(SUM(billing_amount), 2) AS total_revenue,
    SUM(los_days) AS total_los_days,
    ROUND(SUM(billing_amount) / NULLIF(SUM(los_days), 0), 2) AS revenue_per_los_day
FROM patient_admissions
WHERE los_days > 0
GROUP BY department
ORDER BY revenue_per_los_day DESC;

-- 5.5 Insurance type revenue share
SELECT
    insurance_type,
    COUNT(*)                                                        AS patients,
    ROUND(SUM(billing_amount), 2)                                   AS total_billed,
    ROUND(SUM(billing_amount)*100.0/(SELECT SUM(billing_amount) FROM patient_admissions), 2) AS revenue_pct
FROM patient_admissions
GROUP BY insurance_type
ORDER BY total_billed DESC;


-- ============================================================
-- SECTION 6: DOCTOR & STAFF ANALYSIS
-- ============================================================

-- 6.1 Top doctors by patient volume
SELECT
    doctor_id,
    doctor_name,
    COUNT(*)                            AS patients_handled,
    COUNT(DISTINCT department)          AS departments,
    ROUND(AVG(billing_amount), 2)       AS avg_billing,
    ROUND(AVG(los_days), 2)             AS avg_los
FROM patient_admissions
GROUP BY doctor_id, doctor_name
ORDER BY patients_handled DESC;

-- 6.2 Doctor performance: readmission & billing
SELECT
    doctor_name,
    COUNT(*)                                                            AS total_patients,
    SUM(CASE WHEN readmission_flag = 'Yes' THEN 1 ELSE 0 END)          AS readmissions,
    ROUND(SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS readmission_rate,
    ROUND(SUM(billing_amount), 2)                                       AS total_revenue
FROM patient_admissions
GROUP BY doctor_name
ORDER BY total_revenue DESC;

-- 6.3 Staff distribution by role
SELECT
    role,
    COUNT(*)                            AS staff_count,
    ROUND(AVG(monthly_salary_inr), 2)   AS avg_salary,
    ROUND(AVG(overtime_hours), 2)       AS avg_overtime_hrs,
    ROUND(AVG(days_worked), 1)          AS avg_days_worked
FROM staff_schedule
GROUP BY role
ORDER BY staff_count DESC;

-- 6.4 Staff overtime analysis by department
SELECT
    department,
    COUNT(*)                            AS total_staff,
    ROUND(SUM(overtime_hours), 1)       AS total_overtime_hrs,
    ROUND(AVG(overtime_hours), 2)       AS avg_overtime_hrs,
    ROUND(SUM(monthly_salary_inr), 2)   AS total_monthly_payroll
FROM staff_schedule
GROUP BY department
ORDER BY total_overtime_hrs DESC;

-- 6.5 Shift distribution
SELECT
    shift,
    role,
    COUNT(*) AS staff_count
FROM staff_schedule
GROUP BY shift, role
ORDER BY shift, staff_count DESC;


-- ============================================================
-- SECTION 7: READMISSION & RISK ANALYSIS
-- ============================================================

-- 7.1 Readmission rate by department
SELECT
    department,
    COUNT(*)                                                                AS total_patients,
    SUM(CASE WHEN readmission_flag = 'Yes' THEN 1 ELSE 0 END)              AS readmissions,
    ROUND(SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*), 2) AS readmission_rate_pct
FROM patient_admissions
GROUP BY department
ORDER BY readmission_rate_pct DESC;

-- 7.2 Readmission by age group
SELECT
    CASE
        WHEN age < 18 THEN 'Pediatric'
        WHEN age < 35 THEN 'Young Adult'
        WHEN age < 55 THEN 'Middle Age'
        WHEN age < 70 THEN 'Senior'
        ELSE               'Elderly'
    END AS age_group,
    COUNT(*) AS total,
    SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END) AS readmitted,
    ROUND(SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS readmit_pct
FROM patient_admissions
GROUP BY CASE
        WHEN age < 18 THEN 'Pediatric'
        WHEN age < 35 THEN 'Young Adult'
        WHEN age < 55 THEN 'Middle Age'
        WHEN age < 70 THEN 'Senior'
        ELSE               'Elderly'
    END 
ORDER BY readmit_pct DESC;

-- 7.3 High-risk diagnosis: readmission + long LOS
SELECT TOP 15
    diagnosis,
    COUNT(*) AS cases,
    ROUND(AVG(los_days), 2) AS avg_los,
    SUM(CASE WHEN readmission_flag = 'Yes' THEN 1 ELSE 0 END) AS readmissions,
    ROUND(
        SUM(CASE WHEN readmission_flag = 'Yes' THEN 1 ELSE 0 END) * 100.0 / COUNT(*),
        2
    ) AS readmit_rate,
    ROUND(AVG(billing_amount), 2) AS avg_billing
FROM patient_admissions
GROUP BY diagnosis
HAVING COUNT(*) >= 10
ORDER BY readmit_rate DESC;

-- 7.4 Monthly readmission trend
SELECT
    FORMAT(admission_date, 'yyyy-MM') AS [month],
    COUNT(*) AS total_admissions,
    SUM(CASE WHEN readmission_flag = 'Yes' THEN 1 ELSE 0 END) AS readmissions,
    ROUND(
        SUM(CASE WHEN readmission_flag = 'Yes' THEN 1 ELSE 0 END) * 100.0 / COUNT(*),
        2
    ) AS readmit_pct
FROM patient_admissions
GROUP BY FORMAT(admission_date, 'yyyy-MM')
ORDER BY [month];


-- ============================================================
-- SECTION 8: WINDOW FUNCTIONS
-- ============================================================

-- 8.1 Rank departments by monthly revenue (within each month)
SELECT
    month, department, revenue_inr,
    RANK() OVER (PARTITION BY month ORDER BY revenue_inr DESC)        AS revenue_rank,
    DENSE_RANK() OVER (PARTITION BY month ORDER BY revenue_inr DESC)  AS dense_rank
FROM department_kpis
ORDER BY month, revenue_rank;

-- 8.2 Running total revenue per department over time
SELECT
    month, department, revenue_inr,
    SUM(revenue_inr) OVER (PARTITION BY department ORDER BY month
                           ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS cumulative_revenue
FROM department_kpis
ORDER BY department, month;

-- 8.3 Month-over-month revenue change per department
SELECT
    month, department, revenue_inr,
    LAG(revenue_inr)  OVER (PARTITION BY department ORDER BY month) AS prev_month_revenue,
    ROUND(revenue_inr - LAG(revenue_inr) OVER (PARTITION BY department ORDER BY month), 2) AS mom_change,
    ROUND((revenue_inr - LAG(revenue_inr) OVER (PARTITION BY department ORDER BY month))
          / LAG(revenue_inr) OVER (PARTITION BY department ORDER BY month) * 100, 2) AS mom_growth_pct
FROM department_kpis
ORDER BY department, month;

-- 8.4 3-month rolling average admissions per department
SELECT
    month, department, total_admissions,
    ROUND(AVG(total_admissions) OVER (
        PARTITION BY department
        ORDER BY month
        ROWS BETWEEN 2 PRECEDING AND CURRENT ROW
    ), 2) AS rolling_3m_avg
FROM department_kpis
ORDER BY department, month;

-- 8.5 Patient billing percentile within department
SELECT
    patient_id, patient_name, department, billing_amount,
    NTILE(4) OVER (PARTITION BY department ORDER BY billing_amount) AS billing_quartile,
    PERCENT_RANK() OVER (PARTITION BY department ORDER BY billing_amount) AS billing_percentile
FROM patient_admissions
ORDER BY department, billing_amount DESC;

-- 8.6 Top patient by billing in each department (ROW_NUMBER)
SELECT * FROM (
    SELECT
        patient_id, patient_name, department, billing_amount, diagnosis,
        ROW_NUMBER() OVER (PARTITION BY department ORDER BY billing_amount DESC) AS rn
    FROM patient_admissions
) ranked
WHERE rn = 1
ORDER BY billing_amount DESC;

-- 8.7 Department's next month revenue forecast signal (LEAD)
SELECT
    month, department, revenue_inr,
    LEAD(revenue_inr) OVER (PARTITION BY department ORDER BY month) AS next_month_revenue
FROM department_kpis
ORDER BY department, month;


-- ============================================================
-- SECTION 9: CTEs & SUBQUERIES
-- ============================================================

-- 9.1 CTE: Departments above average occupancy rate
WITH dept_occupancy AS (
    SELECT
        department,
        ROUND(AVG(bed_occupancy_rate_pct), 2) AS avg_occupancy
    FROM department_kpis
    GROUP BY department
),
overall_avg AS (
    SELECT ROUND(AVG(bed_occupancy_rate_pct), 2) AS overall_avg
    FROM department_kpis
)
SELECT
    d.department,
    d.avg_occupancy,
    o.overall_avg,
    ROUND(d.avg_occupancy - o.overall_avg, 2) AS above_avg_by
FROM dept_occupancy d
CROSS JOIN overall_avg o
WHERE d.avg_occupancy > o.overall_avg
ORDER BY d.avg_occupancy DESC;

-- 9.2 CTE: High-risk patients (elderly + readmission + high billing)
WITH patient_risk AS (
    SELECT *,
        CASE
            WHEN age >= 70 AND readmission_flag='Yes' AND billing_amount > 200000 THEN 'High Risk'
            WHEN age >= 55 AND readmission_flag='Yes' THEN 'Medium Risk'
            ELSE 'Low Risk'
        END AS risk_category
    FROM patient_admissions
)
SELECT
    risk_category,
    COUNT(*) AS patients,
    ROUND(AVG(billing_amount),2) AS avg_billing,
    ROUND(AVG(los_days),2) AS avg_los,
    STRING_AGG(department, ', ') AS departments
FROM patient_risk
GROUP BY risk_category
ORDER BY COUNT(*) DESC;

-- 9.3 CTE: Department efficiency score
WITH dept_metrics AS (
    SELECT
        department,
        AVG(patient_satisfaction) AS avg_sat,
        AVG(bed_occupancy_rate_pct) AS avg_occ,
        AVG(readmission_rate_pct) AS avg_readmit,
        AVG(mortality_rate_pct) AS avg_mortality,
        AVG(revenue_inr - operating_cost_inr) AS avg_profit
    FROM department_kpis
    GROUP BY department
),
scored AS (
    SELECT *,
        ROUND(
            (avg_sat * 10) +
            (avg_occ * 0.5) -
            (avg_readmit * 2) -
            (avg_mortality * 5) +
            (avg_profit / 100000),
            2
        ) AS efficiency_score
    FROM dept_metrics
)
SELECT
    department,
    efficiency_score,
    RANK() OVER (ORDER BY efficiency_score DESC) AS efficiency_rank
FROM scored
ORDER BY efficiency_score DESC;

-- 9.4 CTE: Multi-step patient journey summary
WITH admission_summary AS (
    SELECT
        patient_id,
        COUNT(admission_id) AS total_visits,
        MIN(admission_date) AS first_visit,
        MAX(admission_date) AS last_visit,
        ROUND(SUM(billing_amount), 2) AS lifetime_value,
        STRING_AGG(department, ' -> ') AS dept_journey
    FROM patient_admissions
    GROUP BY patient_id
    HAVING COUNT(admission_id) > 1
)
SELECT TOP 20 *
FROM admission_summary
ORDER BY total_visits DESC, lifetime_value DESC;

-- 9.5 Subquery: Departments with readmission rate above hospital average
SELECT
    department,
    ROUND(
        SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END) * 100.0 / COUNT(*),
        2
    ) AS dept_readmit_rate,
    (
        SELECT ROUND(
            SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END) * 100.0 / COUNT(*),
            2
        )
        FROM patient_admissions
    ) AS hospital_avg_readmit_rate
FROM patient_admissions
GROUP BY department
HAVING
    SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END) * 100.0 / COUNT(*)
    >
    (
        SELECT
            SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END) * 100.0 / COUNT(*)
        FROM patient_admissions
    )
ORDER BY dept_readmit_rate DESC;


-- ============================================================
-- SECTION 10: KPI SUMMARY DASHBOARD QUERIES
-- ============================================================

-- 10.1 Hospital-level KPI scorecard
SELECT
    COUNT(DISTINCT patient_id)                                              AS unique_patients,
    COUNT(admission_id)                                                     AS total_admissions,
    ROUND(AVG(los_days), 2)                                                 AS avg_length_of_stay,
    ROUND(SUM(billing_amount), 2)                                           AS total_revenue_inr,
    ROUND(AVG(billing_amount), 2)                                           AS avg_billing_per_admission,
    ROUND(SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*), 2) AS overall_readmission_rate_pct,
    SUM(CASE WHEN status='Active' THEN 1 ELSE 0 END)                        AS current_active_patients,
    COUNT(DISTINCT doctor_id)                                                AS active_doctors
FROM patient_admissions;

-- 10.2 Department leaderboard (combined KPI score)
SELECT
    pa.department,
    COUNT(pa.admission_id)                                                  AS admissions,
    ROUND(AVG(pa.billing_amount),2)                                         AS avg_billing,
    ROUND(AVG(pa.los_days),2)                                               AS avg_los,
    ROUND(SUM(CASE WHEN pa.readmission_flag='Yes' THEN 1 ELSE 0 END)*100.0/COUNT(*),2) AS readmit_rate,
    ROUND(AVG(dk.patient_satisfaction),2)                                   AS avg_satisfaction,
    ROUND(AVG(dk.bed_occupancy_rate_pct),2)                                 AS avg_occupancy,
    ROUND(SUM(dk.revenue_inr - dk.operating_cost_inr),2)                   AS net_profit
FROM patient_admissions pa
LEFT JOIN department_kpis dk ON pa.department = dk.department
GROUP BY pa.department
ORDER BY admissions DESC;

-- 10.3 Year-over-year hospital performance
SELECT
    YEAR(admission_date) AS [year],
    COUNT(*) AS admissions,
    ROUND(SUM(billing_amount),2) AS revenue,
    ROUND(AVG(los_days),2) AS avg_los,
    ROUND(
        SUM(CASE WHEN readmission_flag='Yes' THEN 1 ELSE 0 END) * 100.0 / COUNT(*),
        2
    ) AS readmit_rate
FROM patient_admissions
GROUP BY YEAR(admission_date)
ORDER BY [year];
-- 10.4 Best performing month by revenue
SELECT TOP 12
    FORMAT(admission_date,'yyyy-MM') AS [month],
    ROUND(SUM(billing_amount),2) AS revenue,
    COUNT(*) AS admissions,
    RANK() OVER (ORDER BY SUM(billing_amount) DESC) AS revenue_rank
FROM patient_admissions
GROUP BY FORMAT(admission_date,'yyyy-MM')
ORDER BY revenue DESC;

-- ============================================================
-- END OF PROJECT 01: HEALTHCARE OPERATIONS ANALYTICS
-- ============================================================
