
for i=1:length(wz)
   a=cell2mat(wz(i));
   if size(a,2)~=0%��ȡ��ַ
       true_wz=a;
%        href="http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057585/n3057590/c5149611/content.html">
%        ������������ҵ����Ʒ���棨��280����</a></td>
      direc=regexpi(true_wz,'(?<=href=").*?(?=">)','match');%��ȡ��ַ
      tempdirec=cell2mat(direc);
      
      str1=[tempdirec(1:end-12)];
      [sourcefile, status] =urlread(cell2mat(direc));
      strname=cell2mat(regexpi(sourcefile,'(?<=��·��������������ҵ����Ʒ����).*?(?=����</title>)','match'));%��ȡ����
      direc2=regexpi(sourcefile,'(?<=part).*?(?=.doc)','match');%��ȡ�ļ���ַ
      str2=['part',cell2mat(direc2),'.doc'];
      str_last=[str1,str2];
      
      fullURL=[str_last];%������ַ
      filename=['����������ҵ����Ʒ����',strname,'����.doc']; %������ļ���
      [f,status2]=urlwrite(fullURL,filename);%��������
   end
   
end
% http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/index.html
clear

for i=1:14 
%     i=2;%����������ҵ����Ʒ����183��������
[sourcefile, status] =urlread(['http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/index_1274678_',num2str(i),'.html']);
    temp=regexpi(sourcefile,'(?<=<a href=../../../../../n1146285/n1146352/n3054355/n3057292/n3057308/).*?(?=����������ҵ����Ʒ����)','match');
%     temp2=cell2mat(temp);
      for j=1:length(temp)
          tempstr=cell2mat(temp(j));
          tempstr2=tempstr(1:8);
%           http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/c3576899/content.html
          strr=['http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/',tempstr2,'/content.html'];
          [sourcefile1, status] =urlread(strr);
                direc2=regexpi(sourcefile1,'(?<=part).*?(?=.doc)','match');%��ȡ�ļ���ַ
                strname=cell2mat(regexpi(sourcefile1,'(?<=<title>����������ҵ����Ʒ����).*?(?=����)','match'));%��ȡ����
                fullURL=['http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/',tempstr2,'/part',cell2mat(direc2),'.doc'];
      filename=['����������ҵ����Ʒ����',strname,'����.doc']; %������ļ���
%        <title>����������ҵ����Ʒ����183��������</title>
      [f,status2]=urlwrite(fullURL,filename);%��������
      end
end

clear
for i=22:24 
%     i=2;%����������ҵ����Ʒ����183��������
[sourcefile, status] =urlread(['http://www.ndrc.gov.cn/zcfb/zcfbgg/index_',num2str(i),'.html']);
 
%    expr1='<td align="center">(\d\d\d\d-\d\d-\d\d)</td>';
   expr1='<a href="./\d\d\d\d\d\d/t\d\d\d\d\d\d\d\d_\d\d\d\d\d.html" target="_blank">����������ҵ����Ʒ';
   [datefile, date_tokens]= regexp(sourcefile, expr1, 'match', 'tokens'); 
   for j=1:length(datefile)
      temp=cell2mat(datefile(j));
      str1=temp(11:end-27);
      str=['http://www.ndrc.gov.cn/zcfb/zcfbgg',str1];
      [sourcefile2, status] =urlread(str);
             direc2=cell2mat(regexpi(sourcefile2,'(?<=������<A href=).*?(?=.doc)','match'));%��ȡ�ļ���ַ
             strname=cell2mat(regexpi(sourcefile2,'(?<=����������ҵ����Ʒ����).*?(?=����</FONT>)','match'));
             direc3=direc2(3:end);
             str_last=[str(1:end-21),direc3,'.doc'];
   
      filename=['����������ҵ����Ʒ����',strname,'����.doc']; %������ļ���
%        <title>����������ҵ����Ʒ����183��������</title>
      [f,status2]=urlwrite(str_last,filename);%��������
   end
    %http://www.ndrc.gov.cn/zcfb/zcfbgg/200808/t20080813_230739.html
    %<a href="./200808/t20080813_230739.html" target="_blank">����������ҵ����Ʒ(��
end




clear
try
    Word = actxGetRunningServer('Word.Application');  %����word����
catch
    Word = actxserver('Word.Application');
end;
f = [pwd '\����������ҵ����Ʒ����123����.doc'];

 Word.Visible = 0;    % ��set(Word, 'Visible', 1);   %���ÿɼ�
if exist(f,'file');    %�����ļ����ڵĻ�
    Document = Word.Documents.Open(f);      %����ĵ��Ķ���Document
else                %�������򴴽����
    Document = Word.Documents.Add;      
    Document.SaveAs(f);        %�����ĵ�
end

Selection = Word.Selection;               %������ڴ�
Selection.Start=0;
a=[];
num=Document.Range.end;
ii=0;
while ii<=num
    ii=ii+1;
    a=[a,Selection.text];
    Selection.MoveRight;     %��������ƶ�һ��
end
a=a(1:num)             %ȡ�ı������ݵĲ��֣�Ҳ�����ں���ռ�������ֽڣ�����һ���ȡ��a�ĳ��ȶ����ı����ȵ��������ҡ�

Word.Quit

% fileFolder=fullfile('C:\Users\ASUS\Desktop\����������ҵ���Ʒ����\�½��ļ���\excel');
% dirOutput=dir(fullfile(fileFolder,'*'));
% fileNames={dirOutput.name}';
f = [pwd '\����������ҵ����Ʒ����1200����.xlsx'];
exist(f,'file')


%____________________��1

clear
fileFolder=fullfile('C:\Users\ASUS\Desktop\����������ҵ���Ʒ����\�½��ļ���\test\test');%��ȡ�ļ���
dirOutput=dir(fullfile(fileFolder,'*'));
fileNames={dirOutput.name}';
fileNames=fileNames(3:end);


companyname=[];
for tempname=1:length(fileNames)%�Ȼ�ȡ��˾�������б�
     
  [num,txt,raw]=xlsread(cell2mat(fileNames(tempname)));
  clear txt;
  clear num;
% [num,txt,raw]=xlsread('����������ҵ����Ʒ����277����.xlsx');
       for j=1:length(raw(:,1))
         
           tempnum=cell2mat(raw(j,1));
                   if isnan(tempnum)
                      continue
                   end
                   if isempty(strfind(tempnum,'����'))==0
                       break
                   end
           tempnum2=str2num(tempnum(1));
           if  tempnum2==floor(tempnum2)
                 companyname=[companyname;raw(j,2)];
           end
%             strfind(s1,pattern)
                   
       end
end
 companyname=unique(companyname);%��˾������ȡ���
for tempcom=1:length(companyname)
    company=[];
   name=cell2mat(companyname(tempcom)); 
      for tempname=1:length(fileNames)
            [num,txt,raw]=xlsread(cell2mat(fileNames(tempname)));
  clear txt;
  clear num;
              for templocation=1:length(raw(:,2))
                  if isempty(strfind(cell2mat(raw(templocation,2)),name))==0
                      break
                  end%�õ���Ӧ��˾��λ��
              end
              
                  for  jj=(templocation+1):100000 %��ȡͬһ����˾�����µĸ��ֶ���   
                      if isnan(cell2mat(raw(jj,2)))~=1
                          break
                      end
                      if isempty(strfind(cell2mat(raw(jj,5)),'����'))~=0
                          u=[[raw(templocation,3)],[raw(templocation,4)],[raw(jj,5)],[raw(jj,6)],[raw(4,1)]];
                          company=[company;u];
                      end
                  end
%             tempcom=[cell2mat(raw(templocation,3)),cell2mat(raw(templocation,4))];
              
      end
      xlswrite('C:\Users\ASUS\Desktop\����������ҵ���Ʒ����\�½��ļ���\test\test\haha.xlsx',company,name);
end


%_________���

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
           if  tempnum2==floor(tempnum2)
                 companyname=[companyname;raw(j,2)];
           end
%             strfind(s1,pattern)
                   
       end
end
% xlswrite('C:\Users\ASUS\Desktop\����������ҵ���Ʒ����\�½��ļ���\test\test\haha.xlsx',companyname);
companyname=unique(companyname);
%����˾������table
for jj=1:length(companyname)
   tempa={cell2mat(companyname(jj)),[],[],[]};
   tableweneed=[tableweneed;tempa]; 
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
                %__________________________________________________________
                   if isempty(strfind(comname,'��ԭ'))==0%"��ԭ"��˾�滻
                       comnamebefore=cell2mat(regexpi(cell2mat(raw(l,2)),'(?<=��ԭ).*?(?=��)','match'));
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
                   if isempty(strfind(comname,'(ԭ'))==0%"(ԭ"��˾�滻
                       comnamebefore=cell2mat(regexpi(cell2mat(raw(l,2)),'(?<=(ԭ).*?(?=))','match'));
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
                   if isempty(strfind(comname,'('))==0%"("��˾�滻
                       comnamebefore=cell2mat(regexpi(cell2mat(raw(l,2)),'(?<=().*?(?=))','match'));
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
for lastk1=2:length(tableweneed(:,1))
    for lastk2=2:length(tableweneed(:,1))
        if lastk1==lastk2
            continue
        end
        if %A��B�л���B��A��
            
        end
    end   
end
%���Ҫ����




