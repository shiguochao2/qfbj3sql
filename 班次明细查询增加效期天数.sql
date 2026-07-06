if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GP_QueryRetailGroupDetail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GP_QueryRetailGroupDetail]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--16.3.	组别零售汇总查询
--单据日期，单据编号\流水号、收银员、会员卡号、门店编号、门店全名、单据类型（零售单、零售退货单）、组别、所属分类、商品分类、商品编号、商品名称、基本单位、规格、剂型、生产厂商、商品条码、批号、生产日期、有效期至、检验报告号、仓库编号、仓库全名、数量、单位（计量单位）、单价、合计金额、折扣、折扣单价、折扣金额、立减额、成本价、零售价、优惠、处方号、货位编号、货位全名、业务员、最近供货商
--职员编号、职员全名、促销状态、促销单号、供应商编号、供应商名称、利润、利润率、退货原单号
CREATE PROC GP_QueryRetailGroupDetail
(
	@SMode	ctInt = 0,	--零售小票、零售汇总单
	@PosRec	ctComment = '',
	@BillCode	ctShortStr = '',
	@PRec	ctInt = 0,
	@PFullName	ctName = '',
	@GroupID	ctInt = 0,
	@BRec	ctInt = 0,
	@Gift	ctInt = 0,
	@BgnDate	ctDate,
	@EndDate	ctDate,
	@LoginUserRec	ctInt,
	@Erec int=0,--过滤表体职员 
	@BeRec INT =0,--收银员
	@szAtypeid varchar(2000)='',--收款方式，多选,传入的是 rec,rec,rec
	@szKtypeid varchar(200)='',
	@pClass INT=0,--所属分类
	--以下条件是商品信息的
	@classify INT=0, --abc分类
	@isDx INT=1 ,--是否代销 0 不是 1 是		
	@VipCardNo VARCHAR(100)='',
	@sx varchar(100)=''
)
AS
set nocount on
set TRANSACTION ISOLATION LEVEL READ UNCOMMITTED 
--处理abc分类
declare @abcstr varchar(50)
set @abcstr=''
select @abcstr = fullname from tbCodeInfo where [type]=25 and ID= @classify
declare @CRec	ctint
set @CRec = Substring(@sx,0,CHARINDEX('$$',@sx))
if CHARINDEX('$$',@sx) = 0 and @CRec is null set @CRec = @sx
set @CRec = ISNULL(@CRec,0)
DECLARE @sPTypeid varchar(50)
IF @PRec=0 
 SELECT @sPTypeid=''
 ELSE 
 SELECT @sPTypeid=typeId FROM  dbo.ptype WHERE  Rec = @PRec
--SELECT  * INTO #btype FROM   dbo.fn_BTypeRight(@LoginUserRec)
IF @SMode = 0
BEGIN 
   SELECT CheckReportNo,r.ERec as bille,r.BillID,BillType as billtype02,
	--CAST(PriceType as bit) isGift,
	case when PriceType = 30 then 1 else 0 end isGift, 
	BillDate,r.HandOverDate AuditingDate,r.BillCode,v.CardNo
	,case BillType when 151 then  1511 when 152 then 1521 end as billtype
	,r.Posid AS MRec,r.Posid AS MTypeid,tb.FullName AS GroupName,p.typeId AS PTypeID,ProviderId AS BRec,
	JobCode,JobNumber,OutFactoryDate,validityPeriod,GermJobNumber,b.KRec,GRec,b.SaleEtypeid ERec,e.FullName EName,Discount,
	0 tax,PromotionStatus,cf.BillCode RecipeNo,sPromotionBillNo,0 as TaxTotal1,
	(CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) as qty,
		  (CASE WHEN BillType = 151 THEN b.DiscountTotal ELSE -b.DiscountTotal END) AS DiscountTotal,
		CASE WHEN BillType = 151 THEN b.total  ELSE -b.total END AS Total,b.total/b.AssQty AS Price,RetailPrice,		
		(b.DiscountTotal- dLessTotal)/b.AssQty TaxPrice,
		CASE WHEN BillType = 151 THEN (b.DiscountTotal- dLessTotal)/b.AssQty*AssQty ELSE -(b.DiscountTotal- dLessTotal)/b.AssQty*AssQty END TaxTotal02,
		(CASE WHEN BillType = 151 THEN b.DiscountTotal ELSE -b.DiscountTotal END)
			/case (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) when 0 then 1 else (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) end  AS DiscountPrice,--折扣单价
		case b.Unit when 101 then b.costprice*p.TinyRate else b.costprice  end as costprice, --成本均价
		(case b.Unit when 101 then b.costprice*p.TinyRate else b.costprice  end)*(CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) AS CostTotal,--成本金额
		case when r.BillType = 151 then isnull(b.dLessTotal,0) else  -isnull(b.dLessTotal,0) end dLessTotal,
		CASE WHEN r.BillType = 151 THEN '零售开票' ELSE '零售退货' END BillName
		,CASE WHEN BillType = 151 THEN RetailPrice*Qty-DiscountTotal+dLessTotal ELSE -(RetailPrice*Qty-DiscountTotal+dLessTotal) END as yhtotal
		--,rp.Name AS PayNameList
		--,a.FullName PayNameList
		,dbo.fn_GetPayInfo(r.BillID,@SMode,@LoginUserRec) AS PayNameList
		,CASE WHEN b.ptypeSource = 1 THEN 1 ELSE 0 END  ptypeSource,r.explain AS Comment
		,(CASE WHEN BillType = 151 THEN b.realtotal ELSE -b.realtotal END)
		  /case (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) when 0 then 1 else (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) end AS realprice,
		case when r.BillType = 151 then isnull(b.realtotal,0) else  -isnull(b.realtotal,0) end realtotal,
		case when r.BillType = 151 then isnull(b.MinusTotal,0) else  -isnull(b.MinusTotal,0) end MinusTotal,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END JObTaxPrice,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END*CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END AS JObTaxTotal,b.SNTEXT,case b.ValidityPeriod when null then '' else 
				DATEDIFF(day,r.BillDate,b.ValidityPeriod) end limitday
	FROM dbo.RetailBillIndex r 
		INNER JOIN dbo.RetailBill b ON r.BillID = b.BillID
		INNER JOIN dbo.fn_PTypeRight(@LoginUserRec) p ON PRec = p.Rec AND  p.deleted=0
		INNER JOIN dbo.employee e ON e.REC = r.ERec
		INNER JOIN dbo.fn_BTypeRight(@LoginUserRec) bt ON bt.Rec = ProviderId
		INNER JOIN dbo.fn_KTypeRight(@LoginUserRec) k ON k.Rec = r.KRec
		--LEFT JOIN  RetailBillIndex_PayInfo_Name rp ON r.BillID = rp.BillID
		--INNER JOIN RetailPayInfo pay ON pay.BillID = r.BillID 
		--INNER JOIN dbo.Atype a ON pay.ARec = a.Rec 		
		LEFT JOIN (SELECT ID,CardNo FROM dbo.VipCard) v ON r.Vipid = v.ID
		LEFT JOIN (SELECT * FROM dbo.tbCodeInfo WHERE Type = 47) tb ON r.GroupID = tb.ID
		left join (select BillID,BillCode from rOtherBillindex where BillType=801) cf on b.RxBillId=cf.BillID		
		LEFT JOIN dbo.jobinfo j ON b.GJobID=j.jobid
	WHERE (@PosRec = '' OR EXISTS(SELECT 1 FROM  dbo.GetBillTypeTable(@PosRec) pos where r.Posid = pos.ObjectID ))
	AND (b.PRec = @PRec OR p.typeId LIKE + @sPTypeid+'%')
	AND p.FullName LIKE '%' + @PFullName + '%'
	AND r.BillCode LIKE  +'%'+ @BillCode +'%'
	AND (GroupID = @GroupID OR @GroupID = 0)
	AND (ProviderId = @BRec OR @BRec = 0)
	AND (PromotionStatus <> '赠品' OR @Gift = 1)
    AND BillDate BETWEEN @BgnDate AND @EndDate+1
	and ((@Erec=0)or(b.SaleEtypeid=@Erec))
	and draft <>1
	and (@VipCardNo='' OR v.CardNo LIKE '%'+@VipCardNo+'%')
	AND (@szAtypeid ='' or exists(SELECT  1 FROM  dbo.RetailPayInfo WHERE BillID=r.BillID AND ARec IN (Select ObjectID from dbo.GetBillTypeTable(@szAtypeid))))
	AND (r.ERec=@BeRec or @BeRec=0)
    and ((@pClass=0) or EXISTS(select 1 from PtypeWorkRange pr where p.rec = pr.PRec AND  pr.PtypeClassID=@pClass))	
    and ((p.classify = @abcstr)or(@classify=0))    
    and (b.PtypeSource = @isDx OR @isDx=1)
    AND (@CRec = 0 OR EXISTS(SELECT 1 FROM ptypecustom pc WHERE pc.Prec=p.Rec and pc.Crec=@CRec))
	AND (EXISTS(SELECT * FROM dbo.sysdata WHERE SubName = 'QueryRetailShowRegistration' AND SubValue='1') OR r.Comment not like '【门诊】%')
    UNION ALL 
	SELECT CheckReportNo,r.ERec as bille,r.BillID,BillType as billtype02,
	--CAST(PriceType as bit) isGift,
	case when PriceType = 30 then 1 else 0 end isGift, 
	BillDate,r.HandOverDate AuditingDate,r.BillCode,v.CardNo
	,case BillType when 151 then  1511 when 152 then 1521 end as billtype
	,r.Posid AS MRec,r.Posid AS MTypeid,tb.FullName AS GroupName,p.typeId AS PTypeID,ProviderId AS BRec,
	JobCode,JobNumber,OutFactoryDate,validityPeriod,GermJobNumber,b.KRec,GRec,b.SaleEtypeid ERec,e.FullName EName,b.Discount,
	0 tax,PromotionStatus,cf.BillCode RecipeNo,sPromotionBillNo,0 as TaxTotal1,
	(CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) as qty,
		  (CASE WHEN BillType = 151 THEN b.DiscountTotal ELSE -b.DiscountTotal END) AS DiscountTotal,
		CASE WHEN BillType = 151 THEN b.total  ELSE -b.total END AS Total,b.total/b.AssQty AS Price,RetailPrice,		
		(b.DiscountTotal- dLessTotal)/b.AssQty TaxPrice,
		CASE WHEN BillType = 151 THEN (b.DiscountTotal- dLessTotal)/b.AssQty*AssQty ELSE -(b.DiscountTotal- dLessTotal)/b.AssQty*AssQty END TaxTotal02,
		(CASE WHEN BillType = 151 THEN b.DiscountTotal ELSE -b.DiscountTotal END)
			/case (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) when 0 then 1 else (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) end  AS DiscountPrice,--折扣单价
		case b.Unit when 101 then b.costprice*p.TinyRate else b.costprice  end as costprice, --成本均价
		(case b.Unit when 101 then b.costprice*p.TinyRate else b.costprice  end)*(CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) AS CostTotal,--成本金额
		case when r.BillType = 151 then isnull(b.dLessTotal,0) else  -isnull(b.dLessTotal,0) end dLessTotal,
		CASE WHEN r.BillType = 151 THEN '零售开票' ELSE '零售退货' END BillName
		,CASE WHEN BillType = 151 THEN RetailPrice*Qty-DiscountTotal+dLessTotal ELSE -(RetailPrice*Qty-DiscountTotal+dLessTotal) END as yhtotal
		,rp.Name AS PayNameList
		--,a.FullName PayNameList
		--,dbo.fn_GetPayInfo(r.BillID,@SMode,@LoginUserRec) AS PayNameList
		,CASE WHEN b.ptypeSource = 1 THEN 1 ELSE 0 END  ptypeSource,r.explain AS Comment
		,(CASE WHEN BillType = 151 THEN b.realtotal ELSE -b.realtotal END)
		  /case (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) when 0 then 1 else (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) end AS realprice
		,case when r.BillType = 151 then isnull(b.realtotal,0) else  -isnull(b.realtotal,0) end realtotal,
		case when r.BillType = 151 then isnull(b.MinusTotal,0) else  -isnull(b.MinusTotal,0) end MinusTotal,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END JObTaxPrice,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END*CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END AS JObTaxTotal,b.SNTEXT,case b.ValidityPeriod when null then '' else 
				DATEDIFF(day,r.BillDate,b.ValidityPeriod) end limitday
	FROM dbo.RetailBillIndex_His r 
		INNER JOIN dbo.RetailBill_His b ON r.BillID = b.BillID
		INNER JOIN dbo.fn_PTypeRight(@LoginUserRec) p ON PRec = p.Rec AND  p.deleted=0
		INNER JOIN dbo.employee e ON e.REC = r.ERec
		INNER JOIN dbo.fn_BTypeRight(@LoginUserRec) bt ON bt.Rec = ProviderId
		INNER JOIN dbo.fn_KTypeRight(@LoginUserRec) k ON k.Rec = r.KRec
		LEFT JOIN  RetailBillIndex_PayInfo_Name rp ON r.BillID = rp.BillID
		--INNER JOIN dbo.RetailPayInfo_his pay ON pay.BillID = r.BillID 
		--INNER JOIN dbo.Atype a ON pay.ARec = a.Rec 		
		LEFT JOIN  dbo.VipCard v ON r.Vipid = v.ID
		LEFT JOIN (SELECT * FROM dbo.tbCodeInfo WHERE Type = 47) tb ON r.GroupID = tb.ID
		left join (select BillID,BillCode from rOtherBillindex where BillType=801) cf on b.RxBillId=cf.BillID		
		LEFT JOIN dbo.jobinfo j ON b.GJobID=j.jobid
	WHERE (@PosRec = '' OR EXISTS(SELECT 1 FROM  dbo.GetBillTypeTable(@PosRec) pos where r.Posid = pos.ObjectID ))
	AND (b.PRec = @PRec OR p.typeId LIKE + @sPTypeid+'%')
	AND p.FullName LIKE '%' + @PFullName + '%'
	AND r.BillCode LIKE  +'%'+ @BillCode +'%'
	AND (GroupID = @GroupID OR @GroupID = 0)
	AND (ProviderId = @BRec OR @BRec = 0)
	AND (PromotionStatus <> '赠品' OR @Gift = 1)
      AND BillDate BETWEEN @BgnDate AND @EndDate+1
	and ((@Erec=0)or(b.SaleEtypeid=@Erec))
	and draft <>1
	and (@VipCardNo='' OR v.CardNo LIKE '%'+@VipCardNo+'%')
	AND (@szAtypeid ='' or exists(SELECT  1 FROM  dbo.RetailPayInfo_his WHERE BillID=r.BillID AND ARec IN (Select ObjectID from dbo.GetBillTypeTable(@szAtypeid))))
	AND (r.ERec=@BeRec or @BeRec=0)
    and ((@pClass=0) or EXISTS(select 1 from PtypeWorkRange pr where p.rec = pr.PRec AND  pr.PtypeClassID=@pClass))	
    and ((p.classify = @abcstr)or(@classify=0))    
    and (b.PtypeSource = @isDx OR @isDx=1)
    AND (@CRec = 0 OR EXISTS(SELECT 1 FROM ptypecustom pc WHERE pc.Prec=p.Rec and pc.Crec=@CRec))    
	AND (EXISTS(SELECT * FROM dbo.sysdata WHERE SubName = 'QueryRetailShowRegistration' AND SubValue='1') OR r.Comment not like '【门诊】%')
	ORDER BY BillDate
END 
ELSE
	SELECT CheckReportNo,BillDate,r.AuditingDate,BillCode,CardNo,BillType,r.Posid AS MRec,r.Posid AS MTypeid,tb.FullName AS GroupName,p.typeId AS PTypeID,ProviderId AS BRec,
		CASE WHEN BillType = 151 THEN b.Qty ELSE -b.qty END qty,
		b.DiscountPrice * CASE WHEN BillType = 151 THEN b.Qty ELSE -b.qty END AS DiscountTotal,
		CASE WHEN r.BillType = 151 then b.Qty*b.SalePrice else -b.Qty*b.SalePrice end AS Total
		,b.SalePrice AS Price
		,RetailPrice,
		JobCode,JobNumber,OutFactoryDate,validityPeriod,GermJobNumber,b.KRec,GRec ,r.ERec as erec,e.FullName EName,
		TaxPrice
		,CASE WHEN BillType = 151 THEN TaxTotal else -TaxTotal end as TaxTotal02
		,b.tax,CASE WHEN PriceType = 0 THEN '' ELSE '赠品' END as PromotionStatus,
		Discount AS Discount,
		b.DiscountPrice AS DiscountPrice,
		b.costprice as costprice,
		case when r.BillType = 151 then dbo.FormatTotal(b.CostPrice*b.Qty) else -dbo.FormatTotal(b.CostPrice*b.Qty) end AS CostTotal,
		0 dLessTotal,CASE WHEN r.BillType = 151 THEN '零售开票' ELSE '零售退货' END BillName
		,(CASE WHEN BillType = 151 THEN TaxTotal else -TaxTotal end)-(b.DiscountPrice * (CASE WHEN BillType = 151 THEN b.Qty ELSE -b.qty END)) as TaxTotal1,
		dbo.fn_GetPayInfo(r.BillID,@SMode,@LoginUserRec) AS PayNameList,
		--a.FullName PayNameList,
		--rp.Name AS PayNameList,
		CASE WHEN b.ptypeSource = 1 THEN 1 ELSE 0 END ptypeSource
		,(CASE WHEN BillType = 151 THEN b.realtotal ELSE -b.realtotal END)
		  /case (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) when 0 then 1 else (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) end AS realprice
		,case when r.BillType = 151 then isnull(b.realtotal,0) else  -isnull(b.realtotal,0) end realtotal
		,case when r.BillType = 151 then isnull(b.TaxTotal - b.realtotal,0) else  -isnull(b.TaxTotal - b.realtotal,0) end realdLess,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END JObTaxPrice,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END*CASE WHEN BillType = 151 THEN b.Qty ELSE -b.Qty END AS JObTaxTotal,b.SNTEXT,case b.ValidityPeriod when null then '' else 
				DATEDIFF(day,r.BillDate,b.ValidityPeriod) end limitday
	FROM dbo.BillIndex r INNER JOIN dbo.SaleBill b ON r.BillID = b.BillID
		INNER JOIN dbo.fn_PTypeRight(@LoginUserRec) p ON PRec = p.Rec  AND  p.deleted=0
		INNER JOIN dbo.employee e ON e.REC = r.BillE
		INNER JOIN dbo.fn_BTypeRight(@LoginUserRec) bt ON bt.Rec = ProviderId
		INNER JOIN dbo.fn_KTypeRight(@LoginUserRec) k ON k.Rec = r.KRec
		--LEFT JOIN  BillIndex_PayInfo_Name rp ON r.BillID = rp.BillID
		--INNER JOIN BillPayInfo pay ON pay.BillID = r.BillID 
		--INNER JOIN dbo.Atype a ON pay.ARec = a.Rec 				
		LEFT JOIN (SELECT ID,CardNo FROM dbo.VipCard) v ON r.VipCardID = v.ID
		LEFT JOIN (SELECT * FROM dbo.tbCodeInfo WHERE Type = 47) tb ON r.SendWay = tb.ID
		LEFT JOIN dbo.jobinfo j ON b.GJobID=j.jobid
	WHERE BillType IN (151,152) 
	AND (@PosRec = '' OR EXISTS(SELECT 1 FROM  dbo.GetBillTypeTable(@PosRec) pos where r.Posid = pos.ObjectID ))
	AND (b.PRec = @PRec OR p.typeId LIKE + @sPTypeid+'%')
	AND p.FullName LIKE '%' + @PFullName + '%'
	AND r.BillCode LIKE  +'%'+ @BillCode +'%'
	AND (r.SendWay = CAST(@GroupID AS VARCHAR(10)) OR @GroupID = 0)
	AND (ProviderId = @BRec OR @BRec = 0)
	AND (PriceType = @Gift OR @Gift = 1)
	AND BillDate BETWEEN @BgnDate AND @EndDate	
	and ((@Erec=0)or(b.SaleEtypeid=@Erec))
	AND (@szAtypeid ='' 
	      or exists(SELECT  1 FROM BillPayInfo WHERE BillID=r.BillID AND ARec IN (Select ObjectID from dbo.GetBillTypeTable(@szAtypeid))))
	AND (r.ERec=@BeRec or @BeRec=0)
    and ((@pClass=0) or EXISTS(select 1 from PtypeWorkRange pr where p.rec = pr.PRec AND  pr.PtypeClassID=@pClass))	
    and ((p.classify = @abcstr)or(@classify=0))    
    and (b.PtypeSource = @isDx OR @isDx=1)
    AND (@CRec = 0 OR EXISTS(SELECT 1 FROM ptypecustom pc WHERE pc.Prec=p.Rec and pc.Crec=@CRec))
    and (@VipCardNo='' OR v.CardNo LIKE '%'+@VipCardNo+'%')
    UNION ALL 
SELECT CheckReportNo,BillDate,r.AuditingDate,BillCode,CardNo,BillType,r.Posid AS MRec,r.Posid AS MTypeid,tb.FullName AS GroupName,p.typeId AS PTypeID,ProviderId AS BRec,
		CASE WHEN BillType = 151 THEN b.Qty ELSE -b.qty END qty,
		b.DiscountPrice * CASE WHEN BillType = 151 THEN b.Qty ELSE -b.qty END AS DiscountTotal,
		CASE WHEN r.BillType = 151 then b.Qty*b.SalePrice else -b.Qty*b.SalePrice end AS Total
		,b.SalePrice AS Price
		,RetailPrice,
		JobCode,JobNumber,OutFactoryDate,validityPeriod,GermJobNumber,b.KRec,GRec,r.ERec as erec,e.FullName EName,
		TaxPrice
		,CASE WHEN BillType = 151 THEN TaxTotal else -TaxTotal end as TaxTotal02
		,b.tax,CASE WHEN PriceType = 0 THEN '' ELSE '赠品' END as PromotionStatus,
		Discount AS Discount,
		b.DiscountPrice AS DiscountPrice,
		b.costprice as costprice,
		case when r.BillType = 151 then dbo.FormatTotal(b.CostPrice*b.Qty) else -dbo.FormatTotal(b.CostPrice*b.Qty) end AS CostTotal,
		0 dLessTotal,CASE WHEN r.BillType = 151 THEN '零售开票' ELSE '零售退货' END BillName
		,(CASE WHEN BillType = 151 THEN TaxTotal else -TaxTotal end)-(b.DiscountPrice * (CASE WHEN BillType = 151 THEN b.Qty ELSE -b.qty END)) as TaxTotal1,
		--dbo.fn_GetPayInfo(r.BillID,@SMode,@LoginUserRec) AS PayNameList,
		--a.FullName PayNameList
		rp.Name AS PayNameList
		,CASE WHEN b.ptypeSource = 1 THEN 1 ELSE 0 END ptypeSource
		,(CASE WHEN BillType = 151 THEN b.realtotal ELSE -b.realtotal END)
		  /case (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) when 0 then 1 else (CASE WHEN BillType = 151 THEN b.AssQty ELSE -b.AssQty END) end AS realprice
		,case when r.BillType = 151 then isnull(b.realtotal,0) else  -isnull(b.realtotal,0) end realtotal
		,case when r.BillType = 151 then isnull(b.TaxTotal - b.realtotal,0) else  -isnull(b.TaxTotal - b.realtotal,0) end realdLess,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END JObTaxPrice,
		CASE WHEN j.JObTaxPrice>0 THEN j.JObTaxPrice ELSE b.CostPrice END*CASE WHEN BillType = 151 THEN b.Qty ELSE -b.qty END AS JObTaxTotal,b.SNTEXT,case b.ValidityPeriod when null then '' else 
				DATEDIFF(day,r.BillDate,b.ValidityPeriod) end limitday
	FROM dbo.BillIndex_HIS r INNER JOIN dbo.SaleBill_His b ON r.BillID = b.BillID
		INNER JOIN dbo.fn_PTypeRight(@LoginUserRec) p ON PRec = p.Rec  AND  p.deleted=0
		INNER JOIN dbo.employee e ON e.REC = r.BillE
		INNER JOIN dbo.fn_BTypeRight(@LoginUserRec) bt ON bt.Rec = ProviderId
		INNER JOIN dbo.fn_KTypeRight(@LoginUserRec) k ON k.Rec = r.KRec
		LEFT JOIN  BillIndex_PayInfo_Name rp ON r.BillID = rp.BillID
		--INNER JOIN dbo.BillPayInfo_his pay ON pay.BillID = r.BillID 
		--INNER JOIN dbo.Atype a ON pay.ARec = a.Rec 				
		LEFT JOIN (SELECT ID,CardNo FROM dbo.VipCard) v ON r.VipCardID = v.ID
		LEFT JOIN (SELECT * FROM dbo.tbCodeInfo WHERE Type = 47) tb ON r.SendWay = tb.ID
		LEFT JOIN dbo.jobinfo j ON b.GJobID=j.jobid
	WHERE BillType IN (151,152) 
	AND (@PosRec = '' OR EXISTS(SELECT 1 FROM  dbo.GetBillTypeTable(@PosRec) pos where r.Posid = pos.ObjectID ))
	AND (b.PRec = @PRec OR p.typeId LIKE + @sPTypeid+'%')
	AND p.FullName LIKE '%' + @PFullName + '%'
	AND r.BillCode LIKE  +'%'+ @BillCode +'%'
	AND (r.SendWay = CAST(@GroupID AS VARCHAR(10)) OR @GroupID = 0)
	AND (ProviderId = @BRec OR @BRec = 0)
	AND (PriceType = @Gift OR @Gift = 1)
	AND BillDate BETWEEN @BgnDate AND @EndDate	
	and ((@Erec=0)or(b.SaleEtypeid=@Erec))
	AND (@szAtypeid ='' 
	      or exists(SELECT  1 FROM BillPayInfo_his WHERE BillID=r.BillID AND ARec IN (Select ObjectID from dbo.GetBillTypeTable(@szAtypeid))))
	AND (r.ERec=@BeRec or @BeRec=0)
    and ((@pClass=0) or EXISTS(select 1 from PtypeWorkRange pr where p.rec = pr.PRec AND  pr.PtypeClassID=@pClass))	
    and ((p.classify = @abcstr)or(@classify=0))    
    and (b.PtypeSource = @isDx OR @isDx=1)
    AND (@CRec = 0 OR EXISTS(SELECT 1 FROM ptypecustom pc WHERE pc.Prec=p.Rec and pc.Crec=@CRec))
    and (@VipCardNo='' OR v.CardNo LIKE '%'+@VipCardNo+'%')    
ORDER BY BillDate
--DROP TABLE #btype

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
------------------------------------
delete conQueryInfoList where ModuleID = 921
GO
delete conCondtionList where ModuleID = 921
GO
delete TFieldDefaultSet where ModuleID = 921
GO
delete conBaseInfoSet where ModuleID = 921
GO
delete QueryFilterCondConfig where ModuleID = 921
GO
----零售班次明细查询----
insert into  conQueryInfoList(ModuleID,QueryName,PrintName,ProcName,LimitType,Version,Visible,IsUse,ProcParams,ProcParamsValue,ProcParamsType,DataType,HelpItem,PropValue) values(921,'零售班次明细查询','零售班次明细查询.RWX','GP_QueryRetailGroupDetail',0,255,1,1,'@SMode,@PosRec,@BillCode,@PRec,@PFullName,@GroupID,@BRec,@Gift,@BgnDate,@EndDate,@LoginUserRec,@Erec,@BeRec,@szAtypeid,@szKtypeid,@pClass,@classify,@isDx','','',1,'','ShowConditionFrom=True')
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,34,'@SMode','数据来源',1,'零售小票=0,零售汇总单=1',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,11,'@PosRec','门店',2,'MultSel=true',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,16,'@BillCode','单据编号',3,'',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,1,'@Prec','商品',4,'Service=true;ShowYLXM=true;一类',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,16,'@PFullName','商品过滤',5,'',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,34,'@GroupID','零售班次',6,'PRoc:GP_GettbCodeInfoGroup',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,2,'@BRec','供应商',7,'',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,3,'@BeRec','收银员',8,'isstop=True',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,3,'@Erec','职员',10,'isstop=True',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,34,'@classify','ABC分类',9,'PRoc:z_PtypeABCList',0,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,8,'@szAtypeid','收款方式',11,'szSysFlag=3;SZTYPEID=00001;SZShowAll=skfs;MultSel=true',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,4,'@szKtypeid','仓库',12,'一类;StockUse=-1;IfShowTiny=t',0,0,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,34,'@pClass','所属分类',13,'PRoc:z_lbtype',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,17,'@BgnDate,@EndDate','开始日期,至',14,'',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,26,'@isDx','是否代销',15,'1',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,26,'@Gift','是否包含赠品',19,'1',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,16,'@VipCardNo','会员卡号',14,'',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(921,47,'@sx','自定义分类',4,'ClassID=1',1,1,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'时间','BillDate',17,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'过账时间','AuditingDate',17,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'单据编号','BillCode',16,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'零售单备注','Comment',16,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'','',3,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'收银员','EName',16,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'会员卡号','CardNo',16,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'','MRec',11,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'单据类型','BillType',30,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'零售班次','GroupName',16,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'','',1,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'','',31,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'','',4,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'','',5,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'数量','Qty',25,'',0,255,0,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'效期天数','limitday',25,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'单价','Price',19,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'金额','Total',21,'',0,255,0,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'折扣','Discount',25,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'折扣单价','DiscountPrice',19,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'折扣金额','DiscountTotal',21,'',0,255,0,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'成本价','costprice',19,'',119,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'成本金额','CostTotal',21,'',119,255,0,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'零售价','RetailPrice',19,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'税率','tax',21,'',0,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'含税单价','TaxPrice',21,'',0,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'税额','TaxTotal1',21,'TaxTotal02-DiscountTotal+dLessTotal',0,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'含税金额','TaxTotal',21,'TaxTotal02+dLessTotal',0,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'立减金额','dLessTotal',21,'',0,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'扣减金额','MinusTotal',21,'',147,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'实收单价','realprice',19,'',147,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'实收金额','realtotal',21,'',147,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'供应商','BRec',2,'',85,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'利润','Profit',21,'discounttotal-CostTotal-dLessTotal',121,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'利润率','ProfitRate',39,'Profit/(discounttotal-dLessTotal)*100',121,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'实收利润','realProfit',21,'realtotal-CostTotal',121,255,2,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'实收利润率','realProfitRate',39,'realProfit/realtotal*100',121,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'促销状态','PromotionStatus',16,'',0,255,3,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'处方号','RecipeNo',16,'',0,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'促销单号','sPromotionBillNo',16,'',0,255,4,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'优惠','yhtotal',21,'',0,255,1,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'是否代销','ptypeSource',26,'',0,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'收款方式','PayNameList',16,'',0,255,4,1,1,0,120,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'是否赠品','isGift',26,'',0,255,5,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'billid','billid',24,'',0,255,0,0,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'bille','bille',24,'',0,255,0,0,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'billtype02','billtype02',24,'',0,255,0,0,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'TaxTotal02','TaxTotal02',21,'',0,255,0,0,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'含税成本金额','JObTaxTotal',21,'',119,255,0,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(921,'追溯码','SNTEXT',16,'',0,255,6,1,1,0,60,0)
Go
--------