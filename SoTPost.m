function SoTPost(fname1,fname2)
%{
%SoTPost
DESCRIPTION
For measuring the internal condition of large, irregularly-shaped tree 
parts, this function combines geometric measurements of the cross-sectional
shape and nail positions obtained by any means at the measurement plane
with existing time of flight measurements. 

NOTES
The coordinates describing the cross-sectional shape and nail positions 
must be recorded in the same coordinate system. 

---------------------------------------------------------------------------
INPUTS
fname1: string - filepath for Excel workbook containing geometry
measurements and required metadata

Specifically, the Excel workbook must be formatted as follows:
Sheet1: float - mx2 set of counter-clockwise ordered coordinates (cm) 
describing the cross-sectional shape of the measured tree part, where m >= 
99
Sheet2: float - mx2 set of counter-clockwise ordered coordinates describing 
the nail positions, where m == number of nails

fname2: string - filepath for existing tomogram .pit file containing time
of flight measurements

OUTPUTS
Sheet3: float - new sheet written to Excel workbook containing 99x4 matrix 
showing the counter-clockwise ordered coordinates of the new shape in the 
first and second columns. In the third and fourth columns, the ID numbers 
of the new and original coordinates are displayed. 
text_geo.pit: post-processed PiCUS tomogram file
Post-processed PiCUS tomogram file: a PiCUS tomogram file will be created 
and stored in the same directory as the input tomogram. The file name will 
be the same as the original with 'NEW' appended to the name.
---------------------------------------------------------------------------
%}

A=xlsread(fname1,1); %Sheet1
B=xlsread(fname1,2); %Sheet2

%Translate into QI
B(:,1)=B(:,1)+(3-min(A(:,1)));
B(:,2)=B(:,2)+(3-min(A(:,2)));
A(:,1)=A(:,1)+(3-min(A(:,1)));
A(:,2)=A(:,2)+(3-min(A(:,2)));

%From set of n points, replace nearest neighbor in cartesian space
if size(A,1) < 99
    error('The length of A must be greater than or equal to 99');
elseif size(A,1) > 99
    C=[resamplePolyline(A,99), (1:99)', zeros(99,1)];
else
    C=[A, (1:length(A))', zeros(length(A),1)];
end
for k = 1:size(B,1)
    ids=setdiff(1:99,find(C(:,4)))';
    distances = sqrt((C(ids,1)-B(k,1)).^2+(C(ids,2)-B(k,2)).^2);
    [~,idx] = min(distances);
    C(ids(idx),1:2)=B(k,1:2);
    C(ids(idx),4)=k;
end

P=table(C(:,1),C(:,2),C(:,3),C(:,4),'VariableNames',{'X','Y','MP','NAIL'});
writetable(P,fname1,'Sheet',3);

%Read original .pit file
D=fileread(fname2);

%Extract metadata
E=str2double(regexp(D,'u=(\d*)','once','tokens'));
F=str2double(regexp(D,'(?<!\w\.?)Hoehe=(\d*)','once','tokens'));
G=str2double(regexp(D,'Norden=(\d*)','once','tokens'));
H=regexp(D,'Zeit=\d*/\d*/\d*\s\d*:\d*(:\d*)?\s\w\w','once','match');
try
    H=datetime(H(6:end),'inputFormat','M/d/yyyy h:m:s a');
catch
    H=datetime(H(6:end),'inputFormat','M/d/yyyy h:m a');
end
I=str2double(regexp(D,'KlopfMethode=(\d*)','once','tokens'));
J=regexp(D,'\<0=(\d*)','tokens');
J=cellfun(@(x) str2double(x),cellfun(@(x) cell2mat(x),J,...
    'UniformOutput',false))';
K=regexp(D,strcat('\d*=(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/',...
    '(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/(\d*)/',...
    '(\d*)/(\d*)/'),'tokens');
K=cell2mat(cellfun(@(x) str2double(x),K,'UniformOutput',false)');

%Write PiCUS geometry .pit file
L{1,1}='';
L{2,1}='[Comments]';
L{3,1}='ort1=';
L{4,1}='ort2=';
L{5,1}='ort3=';
L{6,1}='ort4=';
L{7,1}='Baumnr=0';
L{8,1}='Formular=1';
L{9,1}='Baumart=';
L{10,1}='BaumartLatein=';
L{11,1}=strcat('Zeit=',datestr(H,'mm/dd/yyyyHH:MM:SS AM'));
L{12,1}='StammUHoehe=130';
L{13,1}='StammU=';
L{14,1}='Baumhoehe=';
L{15,1}='KronenD=';
L{16,1}='Longitude=';
L{17,1}='Latitude=';
L{18,1}='KronenansatzHoehe=';
L{19,1}='Baumalter=';
L{20,1}='VitalitaetRoloff=0';
L{21,1}='Neigungsrichtung=';
L{22,1}='Neigungswinkel=0';
L{23,1}='Neigungbei=0';
L{24,1}='Bearbeiter1=';
L{25,1}='allg_kommentare1=';
L{26,1}='auftraggeber1=';
L{27,1}='BildDatei1=';
L{28,1}='BildDatei2=';
L{29,1}='';
L{30,1}='[Main]';
L{31,1}='Sensoranzahl=99';
L{32,1}='MiniSensorenanzahl=12';
L{33,1}=strcat('KlopfMethode=',num2str(I));
L{34,1}='Hammer=0';
L(35:133,1)=sprintfc('ModTyp%-d=0',(1:99)');
L{134,1}='ModulVerstaerkung=0';
L{135,1}='SampleanzahlHuellkurve=40'; %Generally 30 or 40
L{136,1}='gr_dm=0'; %Major diameter
L{137,1}='kl_dm=0'; %Minor diameter
L{138,1}='messPunktAbstand=0';
L{139,1}=strcat('u=',num2str(E)); %Girth
L{140,1}=strcat('Norden=',num2str(C(C(:,4)==G,3))); %North at MP
L{141,1}='Pos1='; %Direction of MP 1
L{142,1}=strcat('Hoehe=',num2str(F)); %Height
L{143,1}='KDhomogen=-1';
L{144,1}='KDLoch=-1';
L{145,1}='KDRiss=-1';
L{146,1}='KDKern=-1';
L{147,1}='KDFaul=-1';
L{148,1}='KDPilz=-1';
L{149,1}='Hauptwindrichtung=1';
L{150,1}='';
L{151,1}='[NagelBorke]';
L{152,1}='NogelBorkeVerwenden=0';
L(153:251,1)=sprintfc('%-d=0/0',(1:99)');
L{252,1}='';
L{253,1}='[TreeSA]';
L{254,1}='Baumart=0';
L{255,1}='Druckfestigkeit=2';
L{256,1}='Luftwiderstandsbeiwert=0.25';
L{257,1}='Durchmesser1=0';
L{258,1}='Durchmesser2=0';
L{259,1}='Rindendicke=1';
L{260,1}='Baumhoehe=1';
L{261,1}='Standort=1';
L{262,1}='Kronenform=1';
L{263,1}='';
L{264,1}='[BPoints]';
L(265:363,1)=compose('%-d=%-f/%-f',[(1:99)',C(:,1:2)]);
L{364,1}='';
L{365,1}='[ZBPoints]';
L(366:464,1)=compose('%-d=%-d',[(1:99)',repmat(F,[99 1])]);
L{465,1}='';
L{466,1}='[MPoints]';
L(467:565,1)=compose('%-d=%-f/%-f',[(1:99)',C(:,1:2)]);
L{566,1}='';
L{567,1}='[ZMPoints]';
L(568:666,1)=compose('%-d=%-d',[(1:99)',repmat(F,[99 1])]);
L{667,1}='';
L{668,1}='[Diagnoses]';
L{669,1}='';
for i=1:99
    if C(i,4) ~= 0
        M=K((length(B)-1)*C(C(:,3)==i,4)-(length(B)-2):(length(B)-1)*C(C(:,3)==i,4),:);
        N=zeros(98,21);
        N(:,1)=C(C(:,3)~=i,3);
        N(C(C(:,3)~=i,4)~=0,2:21)=M;
        L{669+(101*i-100),1}=strcat('[oLink',num2str(i),']');
        L{669+(101*i-99),1}=strcat('0=',num2str(J(C(C(:,3)==i,4),1)));
        L(669+(101*i-98):669+(101*i-1),1)=compose(strcat('%-d=%-d/%-d/',...
            '%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/',...
            '%-d/%-d/%-d/%-d/'),N);
        L{669+(101*i),1}='';
    else %C(i,4) == 0
        N=zeros(98,21);
        N(:,1)=C(C(:,3)~=i,3);
        L{669+(101*i-100),1}=strcat('[oLink',num2str(i),']');
        L{669+(101*i-99),1}='0=0';
        L(669+(101*i-98):669+(101*i-1),1)=compose(strcat('%-d=%-d/%-d/',...
            '%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/%-d/',...
            '%-d/%-d/%-d/%-d/'),N);
        L{669+(101*i),1}='';
    end
end
if ~isempty(regexp(D,'FreqLink','once'))
    O=splitlines(compose('[oMinMaxFreqLink%-d]\n0=%-d\n',...
        [C(C(:,4)~=0,3),J((size(J,1)/3)+1:2*(size(J,1)/3))]));
    L(10669:10669+size(J,1)-1,1)=reshape(O',[size(J,1),1]);
    O=splitlines(compose('[oIntegralFreqLink%-d]\n0=%-d\n',...
        [C(C(:,4)~=0,3),J(2*(size(J,1)/3)+1:end)]));
    L(10669+size(J,1):10669+2*size(J,1)-1,1)=reshape(O',[size(J,1),1]);
end
[fpath,name,~]=fileparts(fname1);
fid=fopen(strcat(fpath,'\',name,' NEW.pit'),'w');
fprintf(fid,'%s\r\n',L{:});
fclose(fid);

end

