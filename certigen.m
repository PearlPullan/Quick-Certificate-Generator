clc 
clear all
close all 
%home 

filename = 'Registration_Details.xls';
[num,txt] = xlsread(filename)
% Reads Excel sheet containing details. Text is read from the file seperately from numbers
len=length(txt)
% obtain length of text in excel i.e the number of certificates to be generated
blankimg = imread('Certificate_Blank.tif');
% Read blank certificate image

for i=1:len
    for j= 2:2 
        text_names(i,j)=txt(i,j)
    end
end
% Obtain names from the txt variable which are in 2nd column

for i=1:len
    for j= 3:3
      text_topic(i,j)=txt(i,j)
    end
end
% Obtain topics from the txt variable which are in 3rd column


%Ignore first row which is heading
for i=2:len
        text_all=[text_names(i,2) text_topic(i,3)]
        % combine names and topics
        
        position = [100 288;120 460];         
        % obtain positions to insert on image, MSPaint or any image editor
        
        RGB = insertText(blankimage,position,text_all,'FontSize',22,'BoxOpacity',0);
        %Provide parameters for function insertText
        %Font size is 22 and opacity of box is 0 means 100% transparent
        
        figure
        imshow(RGB)        
        %Obtain and display figure in color
        
        y=i-1
        filename=['CertificateTopic_' num2str(y)]
        lastf=[filename '.tif']
        saveas(gcf,lastf)
        % generate and save files with .tif extension
end    
