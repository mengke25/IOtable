clear,clc
eta=cell(14,2); 
year=2000:2013
for i=1:14
    eta{i,2}=1999+i
end
[Type,Sheet,Fromat]=xlsfinfo('F:\yida_project\IND_NIOT_nov16.xlsx');
for i=1:length(Sheet)
    [Num,Text,RawData] =xlsread('F:\yida_project\IND_NIOT_nov16.xlsx',Sheet{i});
end
sheet=cell(14,1)
for i=1:14
sheet{i,1}=Sheet{1,i+3}
end
for e=1:14
P=xlsread('F:\yida_project\IND_NIOT_nov16.xlsx',sheet{e,1},'B2:BE57');
Q=xlsread('F:\yida_project\IND_NIOT_nov16.xlsx',sheet{e,1},'B58:BE113');
GO=xlsread('F:\yida_project\IND_NIOT_nov16.xlsx',sheet{e,1},'B121:BE121');
O=(P+Q)./GO
s=size(O);
for i=1:s
    for j =1:s
     K(i,j)=isnan(O(i,j));
     end 
end
  for i=1:s
    for j=1:s
    if K(i,j)==1
        O(i,j)=0;
    end      
   end  
end
I=eye(s);
eta{e,1}=(I-O)^(-1) -I
end
for i=1:14
filetitle=['F:\yida_project\' 'data' num2str(i) '.xlsx'];
dataTmp=eta{i,1};
xlswrite(filetitle,dataTmp);
end
