% function biaotwo()
clear
fileFolder=fullfile('C:\Users\ASUS\Desktop\车辆生产企业与产品公告\新建文件夹\test\test');%获取文件名
dirOutput=dir(fullfile(fileFolder,'*'));
fileNames={dirOutput.name}';
fileNames=fileNames(3:end);

tableweneed={['公司名'],['生产地址'],['批次'],['冲突']};
companyname=[];
for tempname=1:length(fileNames)%先获取公司的名称列表
    
    [num,txt,raw]=xlsread(cell2mat(fileNames(tempname)));
    clear txt;
    clear num;
    
    for j=1:length(raw(:,1))
        
        tempnum=cell2mat(raw(j,1));
        if isnan(tempnum)
            continue
        end
        if isempty(strfind(tempnum,'二、'))==0
            break
        end
        tempnum2=str2num(tempnum(1));
        if  tempnum2==floor(tempnum2)%&& isempty(strfind(cell2mat(raw(j,2)),'同意'))~=0
            companyname=[companyname;raw(j,2)];
        end
        %             strfind(s1,pattern)
        
    end
end
% xlswrite('C:\Users\ASUS\Desktop\车辆生产企业与产品公告\新建文件夹\test\test\haha.xlsx',companyname);

for u=1:length(companyname)
    uu=cell2mat(companyname(u));
    if isempty(strfind(uu,'同意'))==0
        companyname(u)={'无效'};
    end
    if isnan(uu)==1
        companyname(u)={'无效'};
    end
end
companyname=unique(companyname);
%将公司名放入table
for jj=1:length(companyname)
    tempa={cell2mat(companyname(jj)),[],[],[]};
    tableweneed=[tableweneed;tempa];
end
%替换(→(...妈的智障
for pp=1:length(tableweneed(:,1))
    %     pp=128;
    a=cell2mat(tableweneed(pp,1));
    a(strfind(a,'('))='（';
    a(strfind(a,')'))='）';
end
for tempname=1:length(fileNames)%get the changed addission
    
    [num,txt,raw]=xlsread(cell2mat(fileNames(tempname)));
    clear txt;
    clear num;
    %                     if isempty(strfind(tempnum,'二、'))==0
    %                        break
    %                    end
    %先找到“二、”位置
    for two_location=1:length(raw(:,1))
        tempnum=cell2mat(raw(two_location,1));
        if isempty(strfind(tempnum,'二、'))==0
            break
        end
    end
    %往回找，找所有raw(:,2)中有“同意二字的字符窜”
    for tongyitemp=(two_location-2):-1:1
        if isempty(strfind(cell2mat(raw(tongyitemp,2)),'同意'))==0 && isempty(strfind(cell2mat(raw(tongyitemp,2)),'生产地址'))==0
            direc2=cell2mat(regexpi(cell2mat(raw(tongyitemp,2)),'(?<=：).*?(?=分公司)','match'));
            addis=cell2mat(regexpi(cell2mat(raw(tongyitemp,2)),'(?<=生产地址：).*?(?=。)','match'));
            tempa={[direc2,'分公司'],addis,cell2mat(raw(4,1)),[]};
            tableweneed=[tableweneed;tempa];
        else
            break
        end
    end
    
    %找变更扩展产品的位置
    for biangeng_location=1:length(raw(:,1))
        tempnum=cell2mat(raw(biangeng_location,1));
        if isempty(strfind(tempnum,'变更扩展产品'))==0
            break
        end
    end
    %遍历备注
    for l=(biangeng_location+4):length(raw(:,2))
        kkk=cell2mat(raw(l,1));
        if isnan(kkk(1))==1
            break
        end
        %             l=644
        if isempty(strfind(cell2mat(raw(l,5)),'生产地址变更为'))==0
            comname=cell2mat(raw(l,2));
            if isempty(strfind(comname,'('))==0
                comname(strfind(comname,'('))='（';
            end
            if isempty(strfind(comname,')'))==0
                tempcom=strfind(comname,')');
                if tempcom(end)==length(comname)
                    comname(tempcom(end))='）';
                end
            end
            %__________________________________________________________
            if isempty(strfind(comname,'（原'))==0%"（原"公司替换
                comnamebefore=cell2mat(regexpi(comname,'(?<=（原).*?(?=）)','match'));
                comnamenew=comname(1:(end-length(comnamebefore)-3));
                check=cell2mat(strfind(tableweneed(:,1),comnamebefore));
                if check==1%看新的变更地址在不在列表内，如果在，呢么变更，如果不在则新增
                    for checklocation=1:length(tableweneed(:,1))
                        if isempty(cell2mat(strfind(tableweneed(checklocation,1),comnamebefore)))==0%查找其位置
                            break
                        end
                    end
                    %checklocation为变更的地址
                    tableweneed(checklocation,1)={comnamenew};
                    
                else
                    tempa={comnamenew,[],cell2mat(raw(4,1)),[]};
                    tableweneed=[tableweneed;tempa];
                end
                comname=comnamenew;
            end
            
            %______________________________________________________
            
            %__________________________________________________________
            if isempty(strfind(comname,'公司（'))==0%"沈阳森源艾思特福汽车有限公司（抚顺金信客车有限公司）"公司替换
                comnamebefore=cell2mat(regexpi(comname,'(?<=公司（).*?(?=）)','match'));
                comnamenew=comname(1:(end-length(comnamebefore)-2));
                check=cell2mat(strfind(tableweneed(:,1),comnamebefore));
                if check==1%看新的变更地址在不在列表内，如果在，呢么变更，如果不在则新增
                    for checklocation=1:length(tableweneed(:,1))
                        if isempty(cell2mat(strfind(tableweneed(checklocation,1),comnamebefore)))==0%查找其位置
                            break
                        end
                    end
                    %checklocation为变更的地址
                    tableweneed(checklocation,1)={comnamenew};
                    
                else
                    tempa={comnamenew,[],cell2mat(raw(4,1)),[]};
                    tableweneed=[tableweneed;tempa];
                end
                comname=comnamenew;
            end
            
            %______________________________________________________
            
            
            comadd=cell2mat(regexpi(cell2mat(raw(l,5)),'(?<=生产地址变更为“).*?(?=”)','match'));
            check=cell2mat(strfind(tableweneed(:,1),comname));
            if check==1%看新的变更地址在不在列表内，如果在，呢么变更，如果不在则新增
                for checklocation=1:length(tableweneed(:,1))
                    if isempty(cell2mat(strfind(tableweneed(checklocation,1),comname)))==0%查找其位置
                        break
                    end
                end
                %checklocation为变更的地址
                tableweneed(checklocation,2)={comadd};
                tableweneed(checklocation,3)=raw(4,1);
            else
                tempa={comname,comadd,cell2mat(raw(4,1)),[]};
                tableweneed=[tableweneed;tempa];
            end
        end
        %             strfind(tableweneed(:,1),'长安福特汽车有限公司')
    end
    
end


%——————————————————查重
for lastk1=2:length(tableweneed(:,1))%A
    if isempty(cell2mat(tableweneed(lastk1,2)))==1%若为空，则继续
        continue
    end
    
    for lastk2=2:length(tableweneed(:,1))%B
        if isempty(cell2mat(tableweneed(lastk1,2)))==1%若为空，则继续
            continue
        end
        if lastk1==lastk2
            continue
        end
        %lastk1=133
        if isempty(strfind(cell2mat(tableweneed(lastk1,2)),cell2mat(tableweneed(lastk2,2))))~=1%A在B中或者B在A中
            tempstr=['与',cell2mat(tableweneed(lastk2,3)),'的',cell2mat(tableweneed(lastk2,1)),'地址冲突'];
            tableweneed(lastk1,4)={tempstr};
            %tableweneed(checklocation,2)={comadd}
        end
    end
end
%最后要排序
xlswrite('C:\Users\ASUS\Desktop\车辆生产企业与产品公告\新建文件夹\test\test\haha.xlsx',tableweneed);



