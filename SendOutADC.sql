/*
This query is related to SendOutADC.py
The target is to show availability in our warehouse in Europe and in Asia with one Excel file
But, different customers get different information

Author: Mauricio Montilla
*/

SELECT 
	CAST ( g1.g1pno AS INT ) AS "Pn"
	, g2.g2de1 AS "Description"
	, k1.k1ys01 AS "Supplier Name"
	, gq_05.gqqys - gq_05.gqqyr AS "Avail. 05 (Q)"
	, gq_ADC.gqqys - gq_ADC.gqqyr AS "Avail. ADC (Q)"
	, gq_05.gqqyp AS "In transit to 05 (Q)"
	, gq_ADC.gqqyp AS "In transit to ADC (Q)"
	, g1.g1xrpc || g1.g1xrpw AS "Repl. Code"
	, Co_05."Total Qty. Ord" AS "Qty. Ord to REPLACE_1 from 05 (Q)"
	, Co_ADC."Total Qty. Ord" AS "Qty. Ord to REPLACE_1 from ADC (Q)"

FROM
	rexndcdta.alicg1 AS g1
	LEFT JOIN rexndcdta.alicgq AS gq_05 ON gq_05.gqpno = g1.g1pno AND gq_05.gqstor = '05'
	LEFT JOIN rexndcdta.alicgq AS gq_ADC ON gq_ADC.gqpno = g1.g1pno AND gq_ADC.gqstor = 'ADC'
	LEFT JOIN rexndcdta.alicg2 AS g2 ON g2.g2pno = g1.g1pno
	LEFT JOIN rexndcdta.alicgm AS gm ON gm.gmpno = g1.g1pno AND gm.gmstor = '****' AND gm.gmspri = '1'
	LEFT JOIN rexndcdta.alpuk1 AS k1 ON k1.k1spno = gm.gmspno 
	LEFT JOIN (
		SELECT
			cd.cdpno AS "Pn"
			, SUM ( cd.cdqtd ) AS "Total Qty. Ord"
		
		FROM
			rexndcdta.alcocd AS cd
		
		WHERE
			cd.cdstor = '05'
			AND cd.cdxstd IN ( '2' , '7' )
			AND cd.cdcus IN ( 'REPLACE_2' )

		GROUP BY
			cd.cdpno
	) AS Co_05 ON Co_05."Pn" = g1.g1pno
	LEFT JOIN (
		SELECT
			cd.cdpno AS "Pn"
			, SUM ( cd.cdqtd ) AS "Total Qty. Ord"
		
		FROM
			rexndcdta.alcocd AS cd
		
		WHERE
			cd.cdstor = 'ADC'
			AND cd.cdxstd IN ( '2' , '7' )
			AND cd.cdcus IN ( 'REPLACE_2' )

		GROUP BY
			cd.cdpno
	) AS Co_ADC ON Co_ADC."Pn" = g1.g1pno
	
WHERE
	gm.gmspno IN (
		'17241'
		, '27670'
		, '12831'
		, '21933'
		, '20099'
		, '13605'
	)
	AND ( 
		IFNULL ( gq_05.gqqys , 0 ) 
		+ IFNULL ( gq_05.gqqyr , 0 ) 
		+ IFNULL ( gq_05.gqqyp , 0 ) 
		+ IFNULL ( gq_ADC.gqqys , 0 ) 
		+ IFNULL ( gq_ADC.gqqyr , 0 ) 
		+ IFNULL ( gq_ADC.gqqyp, 0 )
	) != 0
----------
