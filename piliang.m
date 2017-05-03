
for i=1:length(wz)
   a=cell2mat(wz(i));
   if size(a,2)~=0%获取网址
       true_wz=a;
%        href="http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057585/n3057590/c5149611/content.html">
%        ·车辆生产企业及产品公告（第280批）</a></td>
      direc=regexpi(true_wz,'(?<=href=").*?(?=">)','match');%获取网址
      tempdirec=cell2mat(direc);
      
      str1=[tempdirec(1:end-12)];
      [sourcefile, status] =urlread(cell2mat(direc));
      strname=cell2mat(regexpi(sourcefile,'(?<=道路机动车辆生产企业及产品（第).*?(?=批）</title>)','match'));%获取批次
      direc2=regexpi(sourcefile,'(?<=part).*?(?=.doc)','match');%获取文件地址
      str2=['part',cell2mat(direc2),'.doc'];
      str_last=[str1,str2];
      
      fullURL=[str_last];%下载网址
      filename=['车辆生产企业及产品（第',strname,'批）.doc']; %保存的文件名
      [f,status2]=urlwrite(fullURL,filename);%下载命令
   end
   
end
% http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/index.html
clear

for i=1:14 
%     i=2;%车辆生产企业及产品（第183批）公告
[sourcefile, status] =urlread(['http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/index_1274678_',num2str(i),'.html']);
    temp=regexpi(sourcefile,'(?<=<a href=../../../../../n1146285/n1146352/n3054355/n3057292/n3057308/).*?(?=车辆生产企业及产品（第)','match');
%     temp2=cell2mat(temp);
      for j=1:length(temp)
          tempstr=cell2mat(temp(j));
          tempstr2=tempstr(1:8);
%           http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/c3576899/content.html
          strr=['http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/',tempstr2,'/content.html'];
          [sourcefile1, status] =urlread(strr);
                direc2=regexpi(sourcefile1,'(?<=part).*?(?=.doc)','match');%获取文件地址
                strname=cell2mat(regexpi(sourcefile1,'(?<=<title>车辆生产企业及产品（第).*?(?=批）)','match'));%获取批次
                fullURL=['http://www.miit.gov.cn/n1146285/n1146352/n3054355/n3057292/n3057308/',tempstr2,'/part',cell2mat(direc2),'.doc'];
      filename=['车辆生产企业及产品（第',strname,'批）.doc']; %保存的文件名
%        <title>车辆生产企业及产品（第183批）公告</title>
      [f,status2]=urlwrite(fullURL,filename);%下载命令
      end
end

clear
for i=22:24 
%     i=2;%车辆生产企业及产品（第183批）公告
[sourcefile, status] =urlread(['http://www.ndrc.gov.cn/zcfb/zcfbgg/index_',num2str(i),'.html']);
 
%    expr1='<td align="center">(\d\d\d\d-\d\d-\d\d)</td>';
   expr1='<a href="./\d\d\d\d\d\d/t\d\d\d\d\d\d\d\d_\d\d\d\d\d.html" target="_blank">车辆生产企业及产品';
   [datefile, date_tokens]= regexp(sourcefile, expr1, 'match', 'tokens'); 
   for j=1:length(datefile)
      temp=cell2mat(datefile(j));
      str1=temp(11:end-27);
      str=['http://www.ndrc.gov.cn/zcfb/zcfbgg',str1];
      [sourcefile2, status] =urlread(str);
             direc2=cell2mat(regexpi(sourcefile2,'(?<=附件：<A href=).*?(?=.doc)','match'));%获取文件地址
             strname=cell2mat(regexpi(sourcefile2,'(?<=车辆生产企业及产品（第).*?(?=批）</FONT>)','match'));
             direc3=direc2(3:end);
             str_last=[str(1:end-21),direc3,'.doc'];
   
      filename=['车辆生产企业及产品（第',strname,'批）.doc']; %保存的文件名
%        <title>车辆生产企业及产品（第183批）公告</title>
      [f,status2]=urlwrite(str_last,filename);%下载命令
   end
    %http://www.ndrc.gov.cn/zcfb/zcfbgg/200808/t20080813_230739.html
    %<a href="./200808/t20080813_230739.html" target="_blank">车辆生产企业及产品(第
end




clear
try
    Word = actxGetRunningServer('Word.Application');  %启动word引擎
catch
    Word = actxserver('Word.Application');
end;
f = [pwd '\车辆生产企业及产品（第123批）.doc'];

 Word.Visible = 0;    % 或set(Word, 'Visible', 1);   %设置可见
if exist(f,'file');    %测试文件存在的话
    Document = Word.Documents.Open(f);      %获得文档的对象Document
else                %不存在则创建添加
    Document = Word.Documents.Add;      
    Document.SaveAs(f);        %保存文档
end

Selection = Word.Selection;               %光标所在处
Selection.Start=0;
a=[];
num=Document.Range.end;
ii=0;
while ii<=num
    ii=ii+1;
    a=[a,Selection.text];
    Selection.MoveRight;     %光标向右移动一格
end
a=a(1:num)             %取文本有内容的部分，也许由于汉字占有两个字节，所以一般读取后a的长度都是文本长度的两倍左右。

Word.Quit

% fileFolder=fullfile('C:\Users\ASUS\Desktop\车辆生产企业与产品公告\新建文件夹\excel');
% dirOutput=dir(fullfile(fileFolder,'*'));
% fileNames={dirOutput.name}';
f = [pwd '\车辆生产企业及产品（第1200批）.xlsx'];
exist(f,'file')


%____________________表1

clear
fileFolder=fullfile('C:\Users\ASUS\Desktop\车辆生产企业与产品公告\新建文件夹\test\test');%获取文件名
dirOutput=dir(fullfile(fileFolder,'*'));
fileNames={dirOutput.name}';
fileNames=fileNames(3:end);


companyname=[];
for tempname=1:length(fileNames)%先获取公司的名称列表
     
  [num,txt,raw]=xlsread(cell2mat(fileNames(tempname)));
  clear txt;
  clear num;
% [num,txt,raw]=xlsread('车辆生产企业及产品（第277批）.xlsx');
       for j=1:length(raw(:,1))
         
           tempnum=cell2mat(raw(j,1));
                   if isnan(tempnum)
                      continue
                   end
                   if isempty(strfind(tempnum,'二、'))==0
                       break
                   end
           tempnum2=str2num(tempnum(1));
           if  tempnum2==floor(tempnum2)
                 companyname=[companyname;raw(j,2)];
           end
%             strfind(s1,pattern)
                   
       end
end
 companyname=unique(companyname);%公司名字提取完毕
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
                  end%得到对应公司的位置
              end
              
                  for  jj=(templocation+1):100000 %获取同一个公司名称下的各种东西   
                      if isnan(cell2mat(raw(jj,2)))~=1
                          break
                      end
                      if isempty(strfind(cell2mat(raw(jj,5)),'底盘'))~=0
                          u=[[raw(templocation,3)],[raw(templocation,4)],[raw(jj,5)],[raw(jj,6)],[raw(4,1)]];
                          company=[company;u];
                      end
                  end
%             tempcom=[cell2mat(raw(templocation,3)),cell2mat(raw(templocation,4))];
              
      end
      xlswrite('C:\Users\ASUS\Desktop\车辆生产企业与产品公告\新建文件夹\test\test\haha.xlsx',company,name);
end


%_________表二

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
           if  tempnum2==floor(tempnum2)
                 companyname=[companyname;raw(j,2)];
           end
%             strfind(s1,pattern)
                   
       end
end
% xlswrite('C:\Users\ASUS\Desktop\车辆生产企业与产品公告\新建文件夹\test\test\haha.xlsx',companyname);
companyname=unique(companyname);
%将公司名放入table
for jj=1:length(companyname)
   tempa={cell2mat(companyname(jj)),[],[],[]};
   tableweneed=[tableweneed;tempa]; 
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
                %__________________________________________________________
                   if isempty(strfind(comname,'（原'))==0%"（原"公司替换
                       comnamebefore=cell2mat(regexpi(cell2mat(raw(l,2)),'(?<=（原).*?(?=）)','match'));
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
                   if isempty(strfind(comname,'(原'))==0%"(原"公司替换
                       comnamebefore=cell2mat(regexpi(cell2mat(raw(l,2)),'(?<=(原).*?(?=))','match'));
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
                   if isempty(strfind(comname,'('))==0%"("公司替换
                       comnamebefore=cell2mat(regexpi(cell2mat(raw(l,2)),'(?<=().*?(?=))','match'));
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
for lastk1=2:length(tableweneed(:,1))
    for lastk2=2:length(tableweneed(:,1))
        if lastk1==lastk2
            continue
        end
        if %A在B中或者B在A中
            
        end
    end   
end
%最后要排序




