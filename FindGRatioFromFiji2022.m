%Inputs, Assumptions, Outputs, and Errors (Hunter Richardson):
%
%This code asks for a ROI zipfile, a spreadsheet (.csv file) with the
%centroid measurements, and a pixel to micron scaling factor for your image.
%The ROI file can be saved from the ROI Manager, the centroid file can be
%saved from using Measure, and the pixel/micron can be found in the scale
%menu. For more detailed instructions on this, please refer to the protocol.
%
%Assumes ImageJ was used to trace the axons into an ROI manager.
%Assumes the Matlab program ReadImageJROI is in the same Matlab pathway.
%Assumes the ROI and centroid files are in the same folder.
%Assumes that the ROI file is in the order: inner ring of myelin sheath for
%axon 1, outer ring of mylein sheath for axon 1, inner ring for axon 2,
%etc.
%Assumes that the centroid X and Y are either the only or the first 
%measurements in the centroid csv file.
%
%Outputs a figure with the axon caliber in the first row and the g ratio in
%the second row for each axon. This can be saved.
%
%Errors: Will output a message if the Gratio for any of the axons is more
%than 1 (outer was in front of inner)
%If the file chosen produces an error or the scaling factor is not real or 
%negative, it will let you choose a different file/ number to try 3 times 
%before ending the program



clc; 
gratioList = []; %creates open list to hold gratios
axonCaliber = []; %creates open list to hold axon calibers 
    
% section adapted from matlab wiki
% Specify the folder where the files live.
f = msgbox('Please select the folder with the desired files.');
uiwait (f);
dataFolder = uigetdir('','Please select the folder with the desired files.');
% Check to make sure that folder actually exists.  Warn user if it doesn't.
if ~isfolder(dataFolder)
    errorMessage = sprintf('Error: The following folder does not exist:\n%s\nPlease specify a new folder.', dataFolder);
    uiwait(warndlg(errorMessage));
    dataFolder = uigetdir(); % Ask for a new one.
    if dataFolder == 0
         % User clicked Cancel
         return;
    end
end

oldFolder = cd(dataFolder); %sets current path to the wanted folder
uiwait(msgbox('Choose the ROI zip file: ')); %ROI zip with ROI's %waits for confirmation the message was read
[ROIfilename, ROIpathname] = uigetfile ('.zip', 'Choose the ROI zip file: '); %asks user for file
ROIpath = fullfile(ROIpathname, ROIfilename); %turns the components into a full file name
uiwait(msgbox('Choose the centroid excel file: ')); %csv file with centroid measurements, one row per ROI  %waits for confirmation the message was read
[centfilename, centpathname] = uigetfile ('.csv', 'Choose the centroid excel file: '); %asks user for file
centPath = fullfile(centpathname, centfilename); %turns the components into a full file name
pixToMicron = str2double(inputdlg({'Input the pixels/micron scale factor (Ex. 243.7770): '})); %asks user for pixel to micron scale factor, must be a number

counter = 0; %set up variables to count the number of input tries
counter2 = 0;
counter3 = 0;
while counter <= 2 %while loop to assure the inputed ROI path works, gives 3 retries
    try 
        axonROI = ReadImageJROI(ROIpath); %Takes the ROI's in the zip file into Matlab as an array
        break %goes out of the while loop if sucessful
    catch %if there is an error
        uiwait(msgbox('Please make sure the ROI file is correct.'));
        [ROIfilename, ROIpathname] = uigetfile ('.zip', 'Choose the ROI zip file: '); %asks again for file
        if ROIfilename == 0
            % User clicked Cancel
            return;
        end
        ROIpath = fullfile(ROIpathname, ROIfilename); 
        counter = counter + 1;
    end
end
while counter2 <= 2 && counter ~= 3 %while loop to assure the inputed centroid path works, gives 3 retries
    try
        centroids = UploadExcel(centPath); %Takes the list of all of the centroids into Matlab
        break %goes out of the while loop if sucessful
    catch %if there is an error
        uiwait(msgbox('Please make sure the centroid file is correct.'));
        [centfilename, centpathname] = uigetfile ('.csv', 'Choose the centroid excel file: '); %asks again for file
        if centfilename == 0
            % User clicked Cancel
            return;
        end
        centPath = fullfile(centpathname, centfilename);
        counter2 = counter2 + 1;
    end
end
while ~isreal(pixToMicron) == 1 || pixToMicron <= 0 %while loop to make sure the entered number is real and positive
        uiwait(msgbox('Please make sure the pixel/ micron scaling factor is correct. Ex. 243.7770'));
        pixToMicron = str2double(inputdlg({'Input the pixels/micron scale factor (Ex. 243.7770): '}));
        counter3 = counter3 + 1;
        if counter3 > 2
            break
        end
end
%if loop to end program after 3 unsuccessful retries
if counter == 3 || counter2 == 3 || counter3 == 3
    uiwait(msgbox('Please look up the filename or scaling factor and try again.'));
    return
end

%sets up variables for switched inner and outer
axonCounter = 0;
wrongList = [];
for i = 1:2:length(axonROI) %for loop that iterates over odd numbers until it reaches the end of the list of ROI's
    axonCounter = axonCounter + 1;
    inner = axonROI{i}; %Takes the first array and assigns it as the inner myelin tracing
    outer = axonROI{i+1}; %second array is assigned as outer myelin
    inner.mnCoordinates; %Takes the pixel coordinates for this space
    outer.mnCoordinates; %Takes pixel coordinates for this space
    centr = centroids{i, 2:3};
    %changes the coordinates from pixels to microns using the given scaling factor and then
    %makes a list of all of the distances from the centroid to each pixel
    innerRad = sqrt(((inner.mnCoordinates(:,1)/pixToMicron)-centr(1)).^2 + ((inner.mnCoordinates(:,2)/pixToMicron)-centr(2)).^2);
    outerRad = sqrt(((outer.mnCoordinates(:,1)/pixToMicron)-centr(1)).^2 + ((outer.mnCoordinates(:,2)/pixToMicron)-centr(2)).^2);
    %takes the avg of the radians
    innerAvgRad = mean(innerRad);
    outerAvgRad = mean(outerRad);
    gratio = innerAvgRad/outerAvgRad; %finds the gratio using avg radians of the circle
    if gratio > 1.0 %if the inner and outer have been switched, it adds the axon number to a list
        wrongList(end+1) = axonCounter; %makes a list of which axons have gratio's more than 1
        axonCaliber (end+1) = innerAvgRad; %adds the axon caliber to a list
        gratioList(end+1) = gratio; %adds the gratio to a list
    else
        axonCaliber (end+1) = innerAvgRad; %adds the axon caliber to a list
        gratioList(end+1) = gratio; %adds the gratio to a list
    end
end


AxonData = [axonCaliber; gratioList]; %makes the two a matrix for displaying as a table
dataTable = array2table(AxonData, 'RowNames',{'Axon Caliber', 'G Ratio'}); %makes the table

%displays the error that some of the axons are switched for inner/ outer myelin
if isempty(wrongList) == 0
    wrongListStr = sprintf('%.0f, ' , wrongList);
    wrongListStr = wrongListStr(1:end-2); %delete last comma and space
    uiwait(msgbox(append('The order of the inner and outer myelin for axon number(s): ', wrongListStr,' should be checked.')));
end

%creates a figure in a separate window
RatioTable = uifigure('Name', 'G Ratios', 'HandleVisibility', 'on');
%shows G Ratio table on the figure
uitable(uigridlayout(RatioTable, [1,1]),'Data',dataTable,'RowName',{'Axon Caliber', 'G Ratio'}); 
figure;
scatter(axonCaliber, gratioList);

%waits before save questions comes up
pause(5);

%ask about saving spreadsheet
answer = uiconfirm(RatioTable, 'Would you like to save this data in a spreadsheet?','Save?', 'Options',{'Yes','Cancel'},'DefaultOption',1,'CancelOption',2);
if answer == "Yes"
    %saves the filename and path given from directory
    [tableFileName, tablePath] = uiputfile({'*.xlsx'; '*.xls'; '*.csv'});
    %saves the table to that file path
    writetable(dataTable, fullfile(tablePath, tableFileName), 'WriteVariableNames', true);
else 
    %user pressed cancel
    return
end
cd(oldFolder);

function excel = UploadExcel(filepath)

clear opts
opts = delimitedTextImportOptions("NumVariables", 6);

% Specify range and delimiter
opts.DataLines = [2, Inf];
opts.Delimiter = ",";

% Specify column names and types
opts.VariableNames = ["VarName1", "X", "Y", "XM", "YM", "Perim"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double"];

% Specify file level properties
opts.ExtraColumnsRule = "ignore";
opts.EmptyLineRule = "read";

% Import the data ignoring warnings
warning('off','all');
excel = readtable(filepath, opts);
warning('on','all');

end