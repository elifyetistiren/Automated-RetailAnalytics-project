SELECT tbShops.shop_name, tbTempResult.Total_Sales, tbTempResult.target_value
                FROM tbShops INNER JOIN
                (SELECT tbTempSales.shop_id, tbTempSales.Total_Sales, tbTempPerf.target_value
                FROM
                    (SELECT shop_id, SUM(sales_price*(1-sales_discount)*sales_quantity) as Total_Sales
                    FROM tbSales
                    WHERE sales_date BETWEEN CDATE(  CDbl(rp_week)  ) AND CDATE(  CDbl(rp_week_end)  )
                    GROUP BY shop_id) as tbTempSales
                INNER Join
                    (SELECT shop_id, target_value
                    FROM tbPerformance
                    WHERE target_week = CDATE(  CDbl(rp_week)  )) as tbTempPerf
                    ON tbTempSales.shop_id = tbTempPerf.shop_id) as tbTempResult
                ON tbShops.shop_id = tbTempResult.shop_id
                ORDER BY tbTempResult.target_value ASC