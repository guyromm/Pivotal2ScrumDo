#!/usr/bin/env python
import xlwt,sys,os,logging as log,csv,re,json,datetime
if os.path.exists('p2sd.json'):
    conf = json.loads(open('p2sd.json','r').read())
else:
    conf = {'loglevel':'WARNING'}
    
log.basicConfig(level=getattr(log,conf['loglevel']))


pfn = sys.argv[1]
opf = sys.argv[2]
operation = sys.argv[3]

coltrans = {'Id':'Story ID'
            ,'Story':'Summary'
            ,'Description':'Detail'
            #,'Estimate':'Points'
            #,'Current State':'Status'
            ,'Owned By':'Assignee'
            ,'Story Type':'Tags'
            ,'URL':'Pivotal URL'
            }
#because scrumdo does not have a representation of some extra pivotal states we'll write them in as labels
statustrans = {
    'unscheduled':'Todo' #?
    ,'unstarted':'Todo'
    ,'started':'Doing'
    ,'finished':'Reviewing' # review by scrum-master
    ,'delivered':'Reviewing' # review by product-owner
    ,'rejected':'Reviewing'
    ,'accepted':'Done'}
statuses_add_labels = ['unscheduled','finished','delivered','rejected']

ignore_fields=['Labels']
assign_fields=['Iteration','Created at','Accepted at','Deadline','Requested By']
scursor = open(pfn,'r')
pivotal_cols=None
iterations={}
csvcursor = csv.reader(scursor, delimiter=',', quotechar='"')
iterations = {} ; iterations_cnt={} ; iteration_dates = {}
tcnt=0
rowcnt=0

import MySQLdb
if 'db_name' in conf:
    db=MySQLdb.connect(passwd=conf['db_passwd'],db=conf['db_name'],port=conf['db_port'],user=conf['db_user'],host=conf['db_host'])
else:
    db=None

for row in csvcursor:
    log.info( 'ROW %s'%rowcnt)
    rowcnt+=1
    if not pivotal_cols:
        pivotal_cols = row
        continue
    idx=0
    orow = {'tasks':[],'comments':[]}
    addlabels=[]
    for scolval in row:
        scolname = pivotal_cols[idx]
        idx+=1
        log.info('%s = %s'%(scolname,scolval))
        if scolname=='Estimate':
            if scolval:
                orow['Points']=int(scolval)
            else:
                orow['Points']=None
        elif scolname=='Current State':
            orow['Status']=scolval #statustrans[scolval]
            if scolval in statuses_add_labels:
                addlabels.append(scolval)
        elif scolname in coltrans:
            orow[coltrans[scolname]]=scolval
        elif scolname in ['Iteration Start','Iteration End']:
            if scolval:
                iidx = pivotal_cols.index('Iteration')
                itid = (row[iidx])
                if itid not in iteration_dates: iteration_dates[itid]={}
                try:
                    dateval = datetime.datetime.strptime(scolval,'%b %d, %Y')
                except:
                    print 'hmm, was in %s = %s (idx=%s)'%(scolname,scolval,idx)
                    raise Exception(row)
                iteration_dates[itid][scolname]=dateval
        elif scolname in ignore_fields: continue
        elif scolname in assign_fields: orow[scolname] = scolval
        elif scolname =='Note':
            if len(scolval):
                try:
                    comment,author,date = re.compile('^(.*) \((.*) - ([^\-]+)\)$',re.M).search(scolval).groups()
                except:
                    raise Exception('cannot extract comment from %s'%scolval)
                orow['comments'].append({'author':author,'date':date,'comment':comment})
        elif scolname=='Task':
            orow['tasks'].append({'task':scolval})
        elif scolname=='Task Status':
            assert 'status' not in orow['tasks'][-1]
            orow['tasks'][-1]['status']=scolval
        else:
            log.warning('what should i do with %s=%s'%(scolname,scolval))
    if orow['Iteration'] not in iterations:
        iterations[orow['Iteration']]=[]
        iterations_cnt[orow['Iteration']]=0
    if len(addlabels):
        orow['Tags']+=','+(','.join(addlabels))

    iterations[orow['Iteration']].append(orow)
    iterations_cnt[orow['Iteration']]+=1
    log.info(orow)
log.info('we ended up with %s iterations (%s)'%(len(iterations),iterations_cnt))

def getuser(c,fullname,notfound):
    assert len(fullname)
    c.execute("select id from auth_user where trim(concat(first_name,' ',last_name))=%s",fullname.strip())
    userids = (c.fetchall())
    if not len(userids):
        log.warning('could not find any users for %s'%fullname)
        notfound['users_times']+=1
        if fullname not in notfound['users']: notfound['users'].append(fullname)
        return None
    assert len(userids)==1
    userid = int(userids[0][0])
    return userid
     
#write rank, expand tasks into detail
for iteration_id,stories in iterations.items():
    rank=0
    for story in stories:
        story['rank']=rank
        story['rank']+=1
        # if len(story['tasks']):
        #     story['Detail']+='\n\n<h3>Tasks</h3>:<br />\n'
        #     for task in story['tasks']:
        #         story['Detail']+='%s - <b>%s</b><br />\n'%(task['task'],task['status'])

if operation=='noop':
    pass
elif operation=='writexls':
    #start writing
    for iteration_id,stories in iterations.items():
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Pivotal export, iteration %s'%iteration_id)
        title_written=False
        y=0
        for story in stories:
            #if story['Points']: raise Exception(story)
            x=0
            if not title_written:
                for k in story:
                    ws.write(y,x,k)
                    x+=1
                title_written=True
                y+=1
                x=0
            for k in story:
                #log.info('about to write %s,%s'%(y,x))
                if type(story[k])==list:
                    wr = json.dumps(story[k])
                elif k=='Status':
                    wr = statustrans[story[k]]
                else:
                    wr = story[k]
                ws.write(y,x,wr)
                x+=1
            y+=1
        if not iteration_id:
            inn = 'backlog'
        else:
            inn = iteration_id
        ofn = '%s_%s.xls'%(opf,inn)
        wb.save(ofn) 
        print('saved %s with %s rows'%(ofn,y))
elif operation=='sprints':

    for itid,itdates in iteration_dates.items():
        print '%s: %s - %s'%(itid,itdates['Iteration Start'],itdates['Iteration End'])
        if db:
            c = db.cursor()    
            log.info('updating iteration %s'%itid)
            res = c.execute("update projects_iteration set start_date=%s,end_date=%s where name=%s",(itdates['Iteration Start'],itdates['Iteration End'],itid))

            #assert res==1,"res = %s while setting iteration %s"%(res,itid)
            log.info('succesfully set iteration dates for %s'%(itid))
elif operation=='printmembers':
    members = {}
    for stories in iterations.values():
        for st in stories:
            for fn in ['Assignee','Requested By']:
                v = st[fn]
                if v not in members and v: members[v]=0
                if v: members[v]+=1
    print 'ITERATION MEMBERS:'
    for member,times in members.items():
        print '%s\t%s'%(member,times)
elif operation=='insextra':
    notfound={'stories_times':0,'users_times':0,'stories':[],'users':[]}
    done={'tasks':0,'assigned':0,'comments':0}
    for iteration_id,stories in iterations.items():
        for story in stories:
            #try to find the story
            c = db.cursor()
            log.info('looking for summary %s'%story['Summary'])
            c.execute("select id from projects_story where summary=%s",story['Summary'])
            storyids = c.fetchall()
            if not len(storyids):
                log.warning('could not find any stories for summary %s'%(story['Summary']))
                notfound['stories_times']+=1
                continue
            assert len(storyids)==1
            storyid = int(storyids[0][0])
            #write tasks
            if len(story['tasks']):
                c.execute("delete from projects_task where story_id=%s",storyid)
                sord = 0
                for task in story['tasks']:
                    if task['status']=='completed':
                        complete=True
                    else:
                        complete=False
                    c.execute("insert into projects_task (story_id,summary,complete,`order`) values(%s,%s,%s,%s)",(storyid,task['task'],complete,sord))
                    sord+=1
                    done['tasks']+=1
            #write creator_id, assignee_id
            for fn,tofn in {'Assignee':'assignee_id','Requested By':'creator_id'}.items():
                if story[fn]:
                    userid = getuser(c,story[fn],notfound)
                    updres = c.execute("update projects_story set "+tofn+"=%s where id=%s",(userid,storyid))
                    done['assigned']+=1
            #insert comments
            if len(story['comments']):
                cres = c.execute("delete from threadedcomments_threadedcomment where object_id=%s",storyid)
                for comment in story['comments']:
                    userid = getuser(c,comment['author'],notfound)
                    if not userid: continue
                    pdate = datetime.datetime.strptime(comment['date'],'%b %d, %Y')
                    log.info('inserted comment with date %s'%comment['date'])                    
                    cres = c.execute("""insert into threadedcomments_threadedcomment (
                    content_type_id
                    ,object_id
                        ,user_id
                        ,date_submitted
                        ,comment
                        ,markup
                        ,is_public,is_approved,date_modified
                       ) values(
                        35 -- content type id
                        ,%s
                        ,%s -- user
                        ,%s -- date submitted
                        ,%s -- comment
                        ,5 -- markup
                        ,1 -- is public
                        ,0 -- is approved
                        ,%s
                        )""",(storyid,userid,pdate,comment['comment'],pdate))
                    done['comments']+=1

                    assert cres==1
    if not notfound['users_times'] and not notfound['stories_times']:
        print 'no orphaned comments'
    else:
        print 'NOT FOUND USERS/STORIES WHILE INSERTING COMMENTS:'
        print notfound
    print 'DONE: %s'%done
else:
    raise Exception('unknown op %s'%operation)
