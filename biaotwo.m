% function biaotwo()
clear
fileFolder=fullfile('C:\Users\ASUS\Desktop\����������ҵ���Ʒ����\�½��ļ���\test\test');%��ȡ�ļ���
dirOutput=dir(fullfile(fileFolder,'*'));
fileNames={dirOutput.name}';
fileNames=fileNames(3:end);

tableweneed={['��˾��'],['������ַ'],['����'],['��ͻ']};
companyname=[];
for tempname=1:length(fileNames)%�Ȼ�ȡ��˾�������б�
    
    [num,txt,raw]=xlsread(cell2mat(fileNames(tempname)));
    clear txt;
    clear num;
    
    for j=1:length(raw(:,1))
        
        tempnum=cell2mat(raw(j,1));
        if isnan(tempnum)
            continue
        end
        if isempty(strfind(tempnum,'����'))==0
            break
        end
        tempnum2=str2num(tempnum(1));
        if  tempnum2==floor(tempnum2)%&& isempty(strfind(cell2mat(raw(j,2)),'ͬ��'))~=0
            companyname=[companyname;raw(j,2)];
        end
        %             strfind(s1,pattern)
        
    end
end
% xlswrite('C:\Users\ASUS\Desktop\����������ҵ���Ʒ����\�½��ļ���\test\test\haha.xlsx',companyname);

for u=1:length(companyname)
    uu=cell2mat(companyname(u));
    if isempty(strfind(uu,'ͬ��'))==0
        companyname(u)={'��Ч'};
    end
    if isnan(uu)==1
        companyname(u)={'��Ч'};
    end
end
companyname=unique(companyname);
%����˾������table
for jj=1:length(companyname)
    tempa={cell2mat(companyname(jj)),[],[],[]};
    tableweneed=[tableweneed;tempa];
end
%�滻(��(...�������
for pp=1:length(tableweneed(:,1))
    %     pp=128;
    a=cell2mat(tableweneed(pp,1));
    a(strfind(a,'('))='��';
    a(strfind(a,')'))='��';
end
for tempname=1:length(fileNames)%get the changed addission
    
    [num,txt,raw]=xlsread(cell2mat(fileNames(tempname)));
    clear txt;
    clear num;
    %                     if isempty(strfind(tempnum,'����'))==0
    %                        break
    %                    end
    %���ҵ���������λ��
    for two_location=1:length(raw(:,1))
        tempnum=cell2mat(raw(two_location,1));
        if isempty(strfind(tempnum,'����'))==0
            break
        end
    end
    %�����ң�������raw(:,2)���С�ͬ����ֵ��ַ��ܡ�
    for tongyitemp=(two_location-2):-1:1
        if isempty(strfind(cell2mat(raw(tongyitemp,2)),'ͬ��'))==0 && isempty(strfind(cell2mat(raw(tongyitemp,2)),'������ַ'))==0
            direc2=cell2mat(regexpi(cell2mat(raw(tongyitemp,2)),'(?<=��).*?(?=�ֹ�˾)','match'));
            addis=cell2mat(regexpi(cell2mat(raw(tongyitemp,2)),'(?<=������ַ��).*?(?=��)','match'));
            tempa={[direc2,'�ֹ�˾'],addis,cell2mat(raw(4,1)),[]};
            tableweneed=[tableweneed;tempa];
        else
            break
        end
    end
    
    %�ұ����չ��Ʒ��λ��
    for biangeng_location=1:length(raw(:,1))
        tempnum=cell2mat(raw(biangeng_location,1));
        if isempty(strfind(tempnum,'�����չ��Ʒ'))==0
            break
        end
    end
    %������ע
    for l=(biangeng_location+4):length(raw(:,2))
        kkk=cell2mat(raw(l,1));
        if isnan(kkk(1))==1
            break
        end
        %             l=644
        if isempty(strfind(cell2mat(raw(l,5)),'������ַ���Ϊ'))==0
            comname=cell2mat(raw(l,2));
            if isempty(strfind(comname,'('))==0
                comname(strfind(comname,'('))='��';
            end
            if isempty(strfind(comname,')'))==0
                tempcom=strfind(comname,')');
                if tempcom(end)==length(comname)
                    comname(tempcom(end))='��';
                end
            end
            %__________________________________________________________
            if isempty(strfind(comname,'��ԭ'))==0%"��ԭ"��˾�滻
                comnamebefore=cell2mat(regexpi(comname,'(?<=��ԭ).*?(?=��)','match'));
                comnamenew=comname(1:(end-length(comnamebefore)-3));
                check=cell2mat(strfind(tableweneed(:,1),comnamebefore));
                if check==1%���µı����ַ�ڲ����б��ڣ�����ڣ���ô������������������
                    for checklocation=1:length(tableweneed(:,1))
                        if isempty(cell2mat(strfind(tableweneed(checklocation,1),comnamebefore)))==0%������λ��
                            break
                        end
                    end
                    %checklocationΪ����ĵ�ַ
                    tableweneed(checklocation,1)={comnamenew};
                    
                else
                    tempa={comnamenew,[],cell2mat(raw(4,1)),[]};
                    tableweneed=[tableweneed;tempa];
                end
                comname=comnamenew;
            end
            
            %______________________________________________________
            
            %__________________________________________________________
            if isempty(strfind(comname,'��˾��'))==0%"����ɭԴ��˼�ظ��������޹�˾����˳���ſͳ����޹�˾��"��˾�滻
                comnamebefore=cell2mat(regexpi(comname,'(?<=��˾��).*?(?=��)','match'));
                comnamenew=comname(1:(end-length(comnamebefore)-2));
                check=cell2mat(strfind(tableweneed(:,1),comnamebefore));
                if check==1%���µı����ַ�ڲ����б��ڣ�����ڣ���ô������������������
                    for checklocation=1:length(tableweneed(:,1))
                        if isempty(cell2mat(strfind(tableweneed(checklocation,1),comnamebefore)))==0%������λ��
                            break
                        end
                    end
                    %checklocationΪ����ĵ�ַ
                    tableweneed(checklocation,1)={comnamenew};
                    
                else
                    tempa={comnamenew,[],cell2mat(raw(4,1)),[]};
                    tableweneed=[tableweneed;tempa];
                end
                comname=comnamenew;
            end
            
            %______________________________________________________
            
            
            comadd=cell2mat(regexpi(cell2mat(raw(l,5)),'(?<=������ַ���Ϊ��).*?(?=��)','match'));
            check=cell2mat(strfind(tableweneed(:,1),comname));
            if check==1%���µı����ַ�ڲ����б��ڣ�����ڣ���ô������������������
                for checklocation=1:length(tableweneed(:,1))
                    if isempty(cell2mat(strfind(tableweneed(checklocation,1),comname)))==0%������λ��
                        break
                    end
                end
                %checklocationΪ����ĵ�ַ
                tableweneed(checklocation,2)={comadd};
                tableweneed(checklocation,3)=raw(4,1);
            else
                tempa={comname,comadd,cell2mat(raw(4,1)),[]};
                tableweneed=[tableweneed;tempa];
            end
        end
        %             strfind(tableweneed(:,1),'���������������޹�˾')
    end
    
end


%����������������������������������������
for lastk1=2:length(tableweneed(:,1))%A
    if isempty(cell2mat(tableweneed(lastk1,2)))==1%��Ϊ�գ������
        continue
    end
    
    for lastk2=2:length(tableweneed(:,1))%B
        if isempty(cell2mat(tableweneed(lastk1,2)))==1%��Ϊ�գ������
            continue
        end
        if lastk1==lastk2
            continue
        end
        %lastk1=133
        if isempty(strfind(cell2mat(tableweneed(lastk1,2)),cell2mat(tableweneed(lastk2,2))))~=1%A��B�л���B��A��
            tempstr=['��',cell2mat(tableweneed(lastk2,3)),'��',cell2mat(tableweneed(lastk2,1)),'��ַ��ͻ'];
            tableweneed(lastk1,4)={tempstr};
            %tableweneed(checklocation,2)={comadd}
        end
    end
end
%���Ҫ����
xlswrite('C:\Users\ASUS\Desktop\����������ҵ���Ʒ����\�½��ļ���\test\test\haha.xlsx',tableweneed);



