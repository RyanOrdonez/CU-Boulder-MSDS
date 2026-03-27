-- SpaceX Falcon 9 SQL Analysis
-- All queries run against SPACEXTABLE in SQLite

-- Task 1: Unique launch sites
SELECT DISTINCT "Launch_Site" FROM SPACEXTABLE;

-- Task 2: 5 records from CCA sites
SELECT * FROM SPACEXTABLE WHERE "Launch_Site" LIKE 'CCA%' LIMIT 5;

-- Task 3: Total payload mass for NASA (CRS)
SELECT SUM("PAYLOAD_MASS__KG_") AS Total_Payload_Mass_kg
FROM SPACEXTABLE WHERE "Customer" = 'NASA (CRS)';

-- Task 4: Average payload mass for F9 v1.1
SELECT AVG("PAYLOAD_MASS__KG_") AS Avg_Payload_Mass_kg
FROM SPACEXTABLE WHERE "Booster_Version" LIKE 'F9 v1.1%';

-- Task 5: First successful ground pad landing
SELECT MIN("Date") AS First_Successful_Ground_Landing
FROM SPACEXTABLE WHERE "Landing_Outcome" = 'Success (ground pad)';

-- Task 6: Drone ship successes with payload 4000-6000 kg
SELECT "Booster_Version" FROM SPACEXTABLE
WHERE "Landing_Outcome" = 'Success (drone ship)'
  AND "PAYLOAD_MASS__KG_" > 4000 AND "PAYLOAD_MASS__KG_" < 6000;

-- Task 7: Mission outcome counts
SELECT "Mission_Outcome", COUNT(*) AS Count
FROM SPACEXTABLE GROUP BY "Mission_Outcome" ORDER BY Count DESC;

-- Task 8: Boosters with maximum payload mass
SELECT "Booster_Version" FROM SPACEXTABLE
WHERE "PAYLOAD_MASS__KG_" = (SELECT MAX("PAYLOAD_MASS__KG_") FROM SPACEXTABLE);

-- Task 9: 2015 failed drone ship landings
SELECT substr("Date", 6, 2) AS Month, "Landing_Outcome",
       "Booster_Version", "Launch_Site"
FROM SPACEXTABLE
WHERE "Landing_Outcome" = 'Failure (drone ship)'
  AND substr("Date", 0, 5) = '2015';

-- Task 10: Landing outcome rankings (2010-06-04 to 2017-03-20)
SELECT "Landing_Outcome", COUNT(*) AS Count
FROM SPACEXTABLE
WHERE "Date" BETWEEN '2010-06-04' AND '2017-03-20'
GROUP BY "Landing_Outcome" ORDER BY Count DESC;
