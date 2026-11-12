-- Practice setup + queries (SQLite)
CREATE TABLE retirees (retiree_id TEXT PRIMARY KEY, name TEXT, age INT, sex TEXT, plan TEXT, region TEXT);
CREATE TABLE claims (claim_id TEXT PRIMARY KEY, retiree_id TEXT, claim_date DATE, claim_type TEXT, cost NUMERIC);
INSERT INTO retirees VALUES ('P001','Ana Mora',67,'F','Advantage','CA');
INSERT INTO retirees VALUES ('P002','Luis Rojas',72,'M','Advantage','CO');
INSERT INTO retirees VALUES ('P003','Maria Vargas',65,'F','Basic','OR');
INSERT INTO retirees VALUES ('P004','Carlos Solano',78,'M','Premium','CA');
INSERT INTO retirees VALUES ('P005','Elena Jimenez',70,'F','Basic','CO');
INSERT INTO retirees VALUES ('P006','Jose Araya',81,'M','Premium','OR');
INSERT INTO retirees VALUES ('P007','Rosa Campos',68,'F','Advantage','CA');
INSERT INTO retirees VALUES ('P008','Pedro Quesada',75,'M','Basic','CA');
INSERT INTO retirees VALUES ('P009','Lucia Mena',66,'F','Premium','CO');
INSERT INTO retirees VALUES ('P010','Mario Chaves',84,'M','Advantage','OR');
INSERT INTO claims VALUES ('C001','P001','2026-01-05','Hospital',4500);
INSERT INTO claims VALUES ('C002','P001','2026-02-10','Pharmacy',320);
INSERT INTO claims VALUES ('C003','P002','2026-01-15','Hospital',12000);
INSERT INTO claims VALUES ('C004','P002','2026-03-01','Specialist',850);
INSERT INTO claims VALUES ('C005','P004','2026-01-20','Hospital',9800);
INSERT INTO claims VALUES ('C006','P004','2026-02-18','Pharmacy',410);
INSERT INTO claims VALUES ('C007','P004','2026-03-22','Specialist',1200);
INSERT INTO claims VALUES ('C008','P005','2026-02-02','Pharmacy',150);
INSERT INTO claims VALUES ('C009','P006','2026-01-09','Hospital',15000);
INSERT INTO claims VALUES ('C010','P006','2026-03-15','Hospital',8700);
INSERT INTO claims VALUES ('C011','P007','2026-02-25','Specialist',600);
INSERT INTO claims VALUES ('C012','P008','2026-03-03','Pharmacy',275);
INSERT INTO claims VALUES ('C013','P010','2026-01-30','Hospital',11000);
INSERT INTO claims VALUES ('C014','P011','2026-02-14','Pharmacy',90);

-- Q1. JOIN: bring plan and age to each claim (= Ex 1)
SELECT c.claim_id, c.retiree_id, r.plan, r.age, c.cost
FROM claims c
LEFT JOIN retirees r ON c.retiree_id = r.retiree_id;

-- Q2. Total cost and # claims per retiree (= Ex 2)
SELECT r.retiree_id,
       COALESCE(SUM(c.cost), 0) AS total_cost,
       COUNT(c.claim_id)        AS num_claims
FROM retirees r
LEFT JOIN claims c ON r.retiree_id = c.retiree_id
GROUP BY r.retiree_id
ORDER BY r.retiree_id;

-- Q3. Summary by plan with HAVING (= Ex 3)
SELECT r.plan,
       COUNT(DISTINCT r.retiree_id)                 AS retirees,
       COALESCE(SUM(c.cost), 0)                     AS total_cost,
       ROUND(SUM(c.cost) * 1.0 / COUNT(DISTINCT r.retiree_id), 2) AS cost_per_retiree,
       ROUND(AVG(c.cost), 2)                        AS avg_cost_per_claim
FROM retirees r
LEFT JOIN claims c ON r.retiree_id = c.retiree_id
GROUP BY r.plan
HAVING SUM(c.cost) > 1000
ORDER BY total_cost DESC;

-- Q4. Summary by claim type (= Ex 4)
SELECT claim_type,
       COUNT(*)  AS num_claims,
       SUM(cost) AS total_cost,
       AVG(cost) AS avg_cost
FROM claims
GROUP BY claim_type
ORDER BY total_cost DESC;

-- Q5. Orphan claims (data quality, = Ex 5)
SELECT c.*
FROM claims c
LEFT JOIN retirees r ON c.retiree_id = r.retiree_id
WHERE r.retiree_id IS NULL;

-- Q6. Retirees with NO claims
SELECT r.retiree_id, r.name, r.plan
FROM retirees r
LEFT JOIN claims c ON r.retiree_id = c.retiree_id
WHERE c.claim_id IS NULL;

-- Q7. FIRST claim per retiree (what VLOOKUP does) with a window function
SELECT retiree_id, claim_id, claim_date, cost
FROM (
    SELECT c.*,
           ROW_NUMBER() OVER (PARTITION BY retiree_id ORDER BY claim_date) AS rn
    FROM claims c
)
WHERE rn = 1;

-- Q8. Running total per retiree (window function)
SELECT retiree_id, claim_date, cost,
       SUM(cost) OVER (PARTITION BY retiree_id ORDER BY claim_date) AS running_cost
FROM claims
ORDER BY retiree_id, claim_date;

-- Q9. The duplicate trap: INNER JOIN multiplies rows
SELECT COUNT(*)                    AS rows_after_join,
       COUNT(DISTINCT r.retiree_id) AS distinct_retirees
FROM retirees r
JOIN claims c ON r.retiree_id = c.retiree_id;
