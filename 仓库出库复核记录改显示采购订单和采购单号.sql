if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[XF_StockOutCheckQuery]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[XF_StockOutCheckQuery]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

Create PROCEDURE XF_StockOutCheckQuery
( 
	@BRec int,
	@ERec int,
	@BillRec int,
	@KRec int,
	@szBeginDate varchar(10),
	@szEndDate varchar(10),
	@szLoginuser varchar(25),
	@szBilltype varchar(100),
	@szPosId int,
	@szRedBill int,
	@Posid INT,
	@ORec	INT=0
)
AS  
declare @LogRec int set @LogRec = dbo.fn_TypeIDToRec('E',@szLoginuser)
declare @ETypeID varchar(25) set @ETypeID = dbo.fn_RecToTypeID('E',@ERec)
declare @KTypeID varchar(25) set @KTypeID = dbo.fn_RecToTypeID('K',@KRec)
declare @BTypeID varchar(25) set @BTypeID = dbo.fn_RecToTypeID('B',@BRec)
declare @BilleTypeID varchar(25) set @BilleTypeID = dbo.fn_RecToTypeID('E',@BillRec)
declare @sql varchar (3000)
declare @sourcebill varchar(2000)
declare @all varchar(1000)
declare @szTempCondition varchar(2000)
select @szTempCondition='a.BRec in (select Rec from #Btypetable) and ((a.BillType = 205 and a.KRec2 in (select Rec from #Ktypetable)) or (a.BillType <> 205 and a.KRec in (select Rec from #Ktypetable))) and'
if @BRec >0 select @szTempCondition=@szTempCondition+' B.typeid like ''' + @BTypeID + '%'' and '
if @ERec >0   select @szTempCondition=@szTempCondition+' ('''+@ETypeID+'''='''' or exists(select *from employee e where e.Rec = a.Erec and e.Typeid like '''+@ETypeID+'%'')) and '
if @BillRec >0    select @szTempCondition=@szTempCondition++' ('''+@BilleTypeID+'''='''' or exists(select *from employee e where e.Rec = a.Bille and e.Typeid like '''+@BilleTypeID+'%'')) and '
if @szBillType  <>'()' select @szTempCondition=@szTempCondition+' a.BillType in (' + '34' + ') and '
if @KRec >0  select @szTempCondition=@szTempCondition+' ((a.BillType = 205 and k2.typeid like ''' + @KTypeID + '%'') or (a.BillType <> 205 and k.Typeid like ''' + @KTypeID + '%'')) and '
if @szRedBill = 0 select  @szTempCondition=@szTempCondition+' redWord = 0 and '
if @Posid <> 0 select @szTempCondition=@szTempCondition+' a.officeid = '+str(@Posid)+' and '
--增加对配送单非加盟店的过滤
if  @szPosId <> 0 
   select @szTempCondition=@szTempCondition+' a.PosId = '+ str(@szPosId)+ ' and '
--select @szTempCondition=@szTempCondition+' ((billtype in (704,707) and a.posid in (select PosId from posinfo where postype =2)) or (billtype in (34,6,209,208,11,45,207,210))) and '
--if @szFilter    <>'' select @szTempCondition=@szTempCondition+@szFilter + ' and '
if @szTempCondition<>'' select @szTempCondition = substring(@szTempCondition,1,len(@szTempCondition)-4)
 
select @sql='
select Rec,typeid,FullName into #Btypetable from dbo.fn_BTypeRight('+Cast(@LogRec as varchar)+')
select rec,typeid into #Ktypetable from dbo.fn_KTypeRight('+Cast(@LogRec as varchar)+')
SELECT a.billID,a.BillCode,a.BillType,a.totalmoney,a.Comment,a.explain,a.Sendway,a.BillDate,a.RedWord,a.BillE,
s.BillID as sBillID,s.shipper,s.checkman,s.upboxman,s.deliver,s.smemo,cast(isnull(s.isout,0) as bit) isout,
	isnull((SELECT FullName
         FROM employee b
         WHERE a.ERec = b.Rec),'''') AS EName,
	isnull((SELECT fullname
         FROM employee b
         WHERE a.Checke = b.Rec),'''') AS checkeName,
	isnull((SELECT fullname
         FROM employee b
         WHERE a.BillE = b.Rec),'''') AS BillEName,
         b.typeid btypeid,B.FullName BName ,a.totalqty,o1.billcode o1billcode
    FROM  vBillIndex_Query a left join OrderIndex o1 on o1.BillID=a.OrderId
inner join fn_OwnerTypeRight('+CAST(@LogRec AS VARCHAR)+') o on a.ORec=o.Rec
left join StockOutCheck s on a.BillID =s.BillID  
left join #Btypetable b on a.BRec = b.Rec 
left join #Ktypetable k on k.rec = a.KRec
left join #Ktypetable k2 on k2.rec = a.KRec2
where a.BillStatus <> 11
and a.billdate >='''+@szBeginDate+''' and a.billdate <='''+@szEndDate+''' 
and a.draft=0 and '+ @szTempCondition+
' AND ('+CAST(@ORec AS VARCHAR)+'=0 OR a.Orec='+CAST(@ORec AS VARCHAR)+')  order by a.billdate,a.BillID'
       --and  b.billstatus<>20  
--print @sql--+@sourcebill+@all
exec (@sql)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
------------------------------------------------------------------
delete conQueryInfoList where ModuleID = 405
GO
delete conCondtionList where ModuleID = 405
GO
delete TFieldDefaultSet where ModuleID = 405
GO
delete conBaseInfoSet where ModuleID = 405
GO
delete QueryFilterCondConfig where ModuleID = 405
GO
----仓库出库复核登记表----
insert into  conQueryInfoList(ModuleID,QueryName,PrintName,ProcName,LimitType,Version,Visible,IsUse,ProcParams,ProcParamsValue,ProcParamsType,DataType,HelpItem,PropValue) values(405,'仓库出库复核登记表','仓库出库复核登记表.RWX','XF_StockOutCheckQuery',0,255,1,1,'@BRec,@ERec,@BillRec,@KRec,@szBeginDate,@szEndDate,@szLoginuser,@szBilltype,@szPosId,@szRedBill,@Posid','0,0,0,0,,','2,3,3,4,17,17,3,16,24,24,24',2,'','')
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(405,2,'@BRec','往来单位',0,'classselect=true;ShowPos=p',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(405,3,'@ERec','经手人',1,'classselect=true;allselect=true',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(405,3,'@BillRec','制单人',2,'classselect=true;allselect=true',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(405,4,'@KRec','仓库',3,'classselect=true;allselect=true;IfShowTiny=t;StockUse=-1',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(405,17,'@szBeginDate,@szEndDate','查询日期,至',5,'',1,1,0)
Go
insert into  conCondtionList(ModuleID,DataType,CondtionName,DisplayName,OrderIndex,OtherInfo,ConditionVisible,ShowVisible,MustField) values(405,48,'@ORec','货主',4,'',1,1,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'单据编号','BillCode',16,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'订单编号','o1BillCode',16,'',0,255,0,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'单据类型','BillType',30,'',0,255,1,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'','',2,'',0,255,2,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'单据金额','totalmoney',21,'',0,255,3,1,1,0,60,1)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'制单人','BillEName',16,'',0,255,4,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'经手人','EName',16,'',0,255,5,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'审核人','checkeName',16,'',0,255,6,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'摘要','Comment',16,'',0,255,7,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'附加说明','explain',16,'',0,255,8,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'送货方式','Sendway',16,'',0,255,9,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'单据过账日期','BillDate',17,'',0,255,10,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'发货人','shipper',16,'',0,255,11,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'复核人','checkman',16,'',0,255,12,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'装箱人','upboxman',16,'',0,255,13,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'送货人','deliver',16,'',0,255,14,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'已出库','isout',26,'',0,255,15,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'备注','smemo',16,'',0,255,16,1,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'红字单据','RedWord',16,'',0,255,17,0,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'BillID','BillID',16,'',0,255,18,0,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'BillE','BillE',16,'',0,255,19,0,1,0,60,0)
Go
insert into  TFieldDefaultSet(ModuleID,DisplayName,FieldName,DataType,Expression,LimitType,Version, OrderIndex,Visible,IsUse,EditFlag,Defaultwidth,ShowTotal) values(405,'合计数量','totalqty',25,'',0,255,3,1,1,0,60,1)
Go