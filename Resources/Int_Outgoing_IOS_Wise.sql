use purple;
select c.partnername as 'SourceNetwork',CallCount AS 'CallCount',ActualDuration AS 'ActualDuration',BilledDuration AS 'BilledDuration' from
(
select supplierid,Count(*) CallCount,sum(durationsec)/60 ActualDuration,
sum(roundedduration)/60 as BilledDuration
from purple.cdrloaded
where calldirection=2
and starttime>=@startTime
and starttime<@endTime
and AnswerTime>=@AnsTime1
and AnswerTime<@AnsTime2
group by supplierid
) x
left join
partner c
on x.supplierid=c.idpartner

