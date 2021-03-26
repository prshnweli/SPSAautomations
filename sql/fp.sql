
SELECT
STU.ID AS 'StudentID',
STU.FN,
STU.LN,
STU.GR AS 'Grade',
(select TE from TCH where stu.CU = TCH.TN and stu.sc = tch.sc and tch.del = 0 ) AS Teacher,
School = ( select LOC.NM from LOC where stu.SC = LOC.CD),
ROUND(GBS.SCR,2) AS 'Score',
CASE ROUND(GBS.SCR, 1)
  WHEN 0.3 THEN 'A'
  WHEN 0.6 THEN 'B'
  WHEN 0.9 THEN 'C'
  WHEN 1.2 THEN 'D'
  WHEN 1.4 THEN 'E'
  WHEN 1.5 THEN 'F'
  WHEN 1.6 THEN 'G'
  WHEN 1.8 THEN 'H'
  WHEN 1.9 THEN 'I'
  WHEN 2.3 THEN 'J'
  WHEN 2.6 THEN 'K'
  WHEN 2.9 THEN 'L'
  WHEN 3.3 THEN 'M'
  WHEN 3.6 THEN 'N'
  WHEN 3.9 THEN 'O'
  WHEN 4.3 THEN 'P'
  WHEN 4.6 THEN 'Q'
  WHEN 4.9 THEN 'R'
  WHEN 5.3 THEN 'S'
  WHEN 5.6 THEN 'T'
  WHEN 5.9 THEN 'U'
  WHEN 6.3 THEN 'V'
  WHEN 6.6 THEN 'W'
  WHEN 6.9 THEN 'X'
  WHEN 7.3 THEN 'Y'
  WHEN 7.6 THEN 'Z'
  ELSE 'No Score'
END AS [Letter]
FROM STU
JOIN GBS
	ON(GBS.SC = STU.SC) AND (GBS.SN = STU.SN)
WHERE GBS.AN = @Gbknum
AND GBS.DC IS NOT NULL
AND STU.SC = @SchoolNum
ORDER BY GR, Teacher, STU.LN, STU.FN
