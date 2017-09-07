use purple;
select c.partnername as 'SourceNetwork',CallCount AS 'CallCount',ActualDuration AS 'ActualDuration',BilledDuration AS 'BilledDuration' from
(
select customerid,Count(*)CallCount ,sum(durationsec)/60 ActualDuration,
sum(RoundedDuration)/60 as BilledDuration
from purple.cdrloaded
where calldirection=3
and starttime>=@startTime
and starttime<@endTime
and AnswerTime>=@AnsTime1
and AnswerTime<@AnsTime2
group by CustomerID
) x
left join
partner c
on x.customerid=c.idpartner;