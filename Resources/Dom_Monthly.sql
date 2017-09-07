use purple;
select c.partnername as 'SourceNetwork',CallCount AS 'CallCount',ActualDuration AS 'ActualDuration',BilledDuration AS 'BilledDuration' from
(
select customerid,supplierid, DATE_FORMAT(date(starttime),'%d/%m/%Y') `Date`,Count(*) as CallCount,sum(durationsec)/60 ActualDuration,
sum(case when (truncate(durationsec-truncate(durationsec,0),1))>=0
then ceiling(durationsec)
else floor(durationsec) end)/60 as RoundedDuration,sum(Duration1)/60 as BilledDuration
from purple.cdrloaded
where calldirection=1
and starttime>=@startTime
and starttime<@endTime
group by Date,CustomerID
) x
left join
partner c
on x.customerid=c.idpartner
left join partner s
on x.supplierid=s.idpartner;
