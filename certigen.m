clc 
clear all
close all 
%home 

filename = 'list.xlsx';
[num,txt] = xlsread(filename);
% Reads data from the first worksheet in the Microsoft Excel spreadsheet file containing details. 
%Text is read from the file seperately from numbers
len=length(txt);
% obtain length of text in excel i.e the number of certificates to be generated
blankimg = imread('jrexcom.png');
% Read blank certificate image

% as first row has headers, we start from second row with index 1
for i=1:len
    
        text_names(i,2)=txt(i,1);
    
end
% Obtain names from the txt variable which are in 2nd column

for i=1:len
    
      text_topic(i,3)=txt(i,2);
    
end
% Obtain team name from the txt variable which are in 3rd column


%Ignore first row which is heading
for i=2:len
        text_all=[text_names(i,2) text_topic(i,3)];
        % combine names and topics
        
        position = [2190 1190;1405 1431];       
        %[ x y pixel coordinates of the start of the first blank; second blank]
        % obtain positions to insert on image, MSPaint by opening the image in
        % it and using the pencil to point to the start pointfor the required text 
        % and note the pixels it shows in the bottom bar
        
        RGB = insertText(blankimg,position,text_all,'FontSize',82,'BoxOpacity',0);
        %Default font with Font size is 22 and opacity of box is 0 means 100% transparent
        
        figure
        imshow(RGB);       
        %Obtain and display figure in color
        
        y=i-1;
        filename=['CertiJR' num2str(y)];
        lastf=[filename '.png'];
        saveas(gcf,lastf);
        % generate and save files with same extension as original img
end    
