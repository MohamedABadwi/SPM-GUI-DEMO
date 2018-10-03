clc;clear;close all

f1 = figure('name','Smart Time Manager','units','norm','pos',[0 0 1 0.97],'menu'...
    ,'none','numbertitle','off');
uicontrol('str','+New','fontsize',12,'units','norm','pos',[0.01 0.85 0.1 0.1],...
    'call',@newtask)
uicontrol('str','Refresh','fontsize',12,'units','norm','pos',[0.12 0.85 0.1 0.1],...
    'call',{@refresh,f1})
refresh(0,0,f1)

function refresh(varargin)
f1 = varargin{3};
[~,~,d] = xlsread('data.csv');
strr = {'Project Name','Number of Tasks','Task Name','Start Date','End Date','Reminder'};
if size(d,1) > 1
   for xx = 1:6
       uicontrol(f1,'str',strr{xx},'style','text','units','norm','pos',[0.01+(xx-1)/10 ...
                0.65 0.1 0.1],'fontsize',16)
       for yy = 1:size(d,1)-1
           uicontrol(f1,'str',num2str(d{yy+1,xx}),'style','text','units','norm','pos',[0.01+(xx-1)/10 ...
               0.65-yy/10 0.1 0.1],'fontsize',16)
       end
   end
end
end
function newtask(varargin)
    f2 = figure('name','New project','units','norm','pos',[0.25 0.25 0.5 0.5],'menu'...
        ,'none','numbertitle','off');
    strs = {'Project Name :','Number of Tasks :'};
    var = {0,0};
    for i = 1:2
        uicontrol(f2,'str',strs{i},'style','text','units','norm','pos',[0.05...
            0.99-i/10 0.23 0.1],'fontsize',12)
        var{i} = uicontrol(f2,'str','','style','edit','units','norm','pos',[0.3 1-i/10 0.3-(2*(i-1)/10) 0.1]...
        ,'fontsize',12);
    end
    uicontrol(f2,'str','Update','units','norm','pos',[0.6 0.8 0.1 0.1]...
        ,'fontsize',12,'call',{@tasksdetails,var,f2})
end

function tasksdetails(varargin)
    f2 = varargin{4};
    num = varargin{3};
    var1 = get(num{1},'str');
    var2 = str2double(get(num{2},'str'));
    strs = {'task name','Start date','End date'};
    var = {0,0};
    varr = {0,0};
    if length(var1) <1
        errordlg('Please Submit the Project Name first','Error msg');
    else
        for i= 1:3
            uicontrol(f2,'str',strs{i},'style','text','units','norm','pos',[0.1+(i-1)/5 ...
                0.65 0.2 0.1],'fontsize',12)
            for x = 1:var2
                var{x,i} = uicontrol(f2,'str','','style','edit','units',...
                    'norm','pos',[0.1+(i-1)/5 0.65-x/10 0.2 0.1],'fontsize',12);
                
            end
        end
        for x = 1:var2
            varr{x} = uicontrol(f2,'str','','style','check','units',...
                    'norm','pos',[0.8 0.65-x/10 0.2 0.1],'fontsize',12);
        end
        uicontrol(f2,'str','Save','units','norm','pos',[0.8 ...
                0.1 0.2 0.1],'fontsize',12,'call',{@savetask,var,var1,var2,varr,f2})
    end
end

function savetask(varargin)
    f2 = varargin{7};
    v1 = varargin{3};
    v2 = varargin{4};
    v3 = varargin{5};
    v4 = varargin{6};
    [~,t] = xlsread('data.csv');
    len = size(t,1);
    %range = sprintf('A%d:D%d',len,len);
    for x = 1:v3
        range = sprintf('A%d:F%d',len+x,len+x);
        data = {v2,v3,get(v1{x,1},'str'),get(v1{x,2},'str'),get(v1{x,3},'str'),...
            get(v4{x},'value')};
        xlswrite('data.csv',data,range);
        Project =v2;
        task = get(v1{x,1},'str');
        location = 'my location';
        Startdate = get(v1{x,2},'str');
        Enddate = get(v1{x,2},'str');
        Remind = get(v4{x},'value');
        delay = 5;
        setappointment(Project,task,location,Startdate,Enddate,Remind,delay)
    end
    close(f2)
    function setappointment(Project,task,location,Startdate,Enddate,Remind,delay)
        h = actxserver('outlook.Application');
        % Create the appointment object
        appointment = h.CreateItem('olAppointmentItem'); 
        % Set some common properties of the appointment object.
        appointment.Subject = Project;
        appointment.Body = ['This appointment for Task',task];
        appointment.Location = location;
        appointment.Start = Startdate;
        appointment.End = Enddate;
        appointment.ReminderSet = Remind;
        appointment.ReminderMinutesBeforeStart = delay;

        %appointment.IsOnlineMeeting = 0
        %Save the appointment
        appointment.Save()
    end
end