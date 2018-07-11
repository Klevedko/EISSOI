package Quartz;

import EISSOI.App;
import org.quartz.*;
import org.quartz.impl.StdSchedulerFactory;

public class CronTriggerExample 
{
    public static void main( String[] args ) throws Exception
    {
    	
    	//JobDetail job = new JobDetail();
    	//job.setName("dummyJobName");
    	//job.setJobClass(HelloJob.class);    	
    	JobDetail job = JobBuilder.newJob(App.class)
				.withIdentity("dummyJobName", "group1").build();

    	//CronTrigger trigger = new CronTrigger();
    	//trigger.setName("dummyTriggerName");
    	//trigger.setCronExpression("0/5 * * * * ?");
    	
    	Trigger trigger = TriggerBuilder
				.newTrigger()
				.withIdentity("dummyTriggerName", "group1")
				.withSchedule(
						CronScheduleBuilder.cronSchedule("0/5 * * * * ?"))
				.build();
    	
    	//schedule it
    	Scheduler scheduler = new StdSchedulerFactory().getScheduler();
    	scheduler.start();
    	scheduler.scheduleJob(job, trigger);
    
    }
}
