var fs = require('fs');
var xlsx = require('node-xlsx').default;
var XLSX = require('xlsx');
//read file
var workbook = XLSX.readFile("E:\\Child Full_Data.xlsx", { cellDates: true }, { raw: false });


//store sheets as variables
var tab1 = workbook.Sheets['Center-Setup'];
var child = workbook.Sheets['child'];
var family = workbook.Sheets['family'];
var t1 = workbook.Sheets['Tuition-1'];
var t2 = workbook.Sheets['Tuition-2'];
var t3 = workbook.Sheets['Tuition-3'];
var t4 = workbook.Sheets['Tuition-4'];
var t5 = workbook.Sheets['Tuition-5'];
var t6 = workbook.Sheets['Tuition-6'];
var t7 = workbook.Sheets['Tuition-7'];
var t8 = workbook.Sheets['Tuition-8'];
var t9 = workbook.Sheets['Tuition-9'];
var t10 = workbook.Sheets['Tuition-10'];
var t11 = workbook.Sheets['Tuition-11'];
var t12 = workbook.Sheets['Tuition-12'];
var waiting = workbook.Sheets["Waiting-copy"];


//convert sheet to array objects
var tab1data = XLSX.utils.sheet_to_json(tab1);
var childdata = XLSX.utils.sheet_to_json(child);
var familydata = XLSX.utils.sheet_to_json(family);
var t1data = XLSX.utils.sheet_to_json(t1);
var t2data = XLSX.utils.sheet_to_json(t2);
var t3data = XLSX.utils.sheet_to_json(t3);
var t4data = XLSX.utils.sheet_to_json(t4);
var t5data = XLSX.utils.sheet_to_json(t5);
var t6data = XLSX.utils.sheet_to_json(t6);
var t7data = XLSX.utils.sheet_to_json(t7);
var t8data = XLSX.utils.sheet_to_json(t8);
var t9data = XLSX.utils.sheet_to_json(t9);
var t10data = XLSX.utils.sheet_to_json(t10);
var t11data = XLSX.utils.sheet_to_json(t11);
var t12data = XLSX.utils.sheet_to_json(t12);
var waitingdata = XLSX.utils.sheet_to_json(waiting);


//tab1
var newtab1data = tab1data.map(function(tab1r) {
    return tab1r;
});


// {s:{c:1, r:4}, e:{c:3, r:18}}   B5:D19
//center.csv
var Crange = { s: { c: 1, r: 4 }, e: { c: 16, r: 5 } }; //B5:Q6
var CdataRange = [];
/* Iterate through each element in the structure */
for (var R = Crange.s.r; R <= Crange.e.r; ++R) {
    let center = [];
    for (var C = Crange.s.c; C <= Crange.e.c; ++C) {
        var cell_address = { c: C, r: R };
        var data = XLSX.utils.encode_cell(cell_address);
        center.push(tab1[data]);
    }

    CdataRange.push(center);
}
var newcentercsv = XLSX.utils.book_new();
var centercsv = XLSX.utils.aoa_to_sheet(CdataRange);

centercsv.A1.v = "fullName";
centercsv.A1.r = "fullName";
centercsv.A1.h = "fullName";
centercsv.A1.w = "fullName";
centercsv.B1.v = "shortName";
centercsv.B1.r = "shortName";
centercsv.B1.h = "shortName";
centercsv.B1.w = "shortName";
centercsv.C1.v = "address";
centercsv.C1.r = "address";
centercsv.C1.h = "address";
centercsv.C1.w = "address";
centercsv.G1.v = "officePhone";
centercsv.G1.r = "officePhone";
centercsv.G1.h = "officePhone";
centercsv.G1.w = "officePhone";

centercsv.H1.v = "supplierNumber"
centercsv.H1.r = "supplierNumber"
centercsv.H1.h = "supplierNumber"
centercsv.H1.w = "supplierNumber"

centercsv.I1.v = "managerName";
centercsv.I1.r = "managerName";
centercsv.I1.h = "managerName";
centercsv.I1.w = "managerName";

centercsv.J1.v = "managerEmail";
centercsv.J1.r = "managerEmail";
centercsv.J1.h = "managerEmail";
centercsv.J1.w = "managerEmail";

centercsv.K1.v = "managerCell";
centercsv.K1.r = "managerCell";
centercsv.K1.h = "managerCell";
centercsv.K1.w = "managerCell";

centercsv.L1.v = "ownerName";
centercsv.L1.r = "ownerName";
centercsv.L1.h = "ownerName";
centercsv.L1.w = "ownerName";

centercsv.M1.v = "ownerEmail";
centercsv.M1.r = "ownerEmail";
centercsv.M1.h = "ownerEmail";
centercsv.M1.w = "ownerEmail";

centercsv.N1.v = "ownerCell";
centercsv.N1.r = "ownerCell";
centercsv.N1.h = "ownerCell";
centercsv.N1.w = "ownerCell";

centercsv.O1.v = "tuitionYear";
centercsv.O1.r = "tuitionYear";
centercsv.O1.h = "tuitionYear";
centercsv.O1.w = "tuitionYear";

centercsv.P1.v = "tuitionMonth";
centercsv.P1.r = "tuitionMonth";
centercsv.P1.h = "tuitionMonth";
centercsv.P1.w = "tuitionMonth";




XLSX.utils.book_append_sheet(newcentercsv, centercsv, "center");
XLSX.writeFile(newcentercsv, "E:\\center.csv");



//program.csv
var Prange = { s: { c: 1, r: 20 }, e: { c: 5, r: 23 } }; //B21 : F24
var PdataRange = [];
/* Iterate through each element in the structure */
for (var R = Prange.s.r; R <= Prange.e.r; ++R) {
    let program = [];
    for (var C = Prange.s.c; C <= Prange.e.c; ++C) {
        var cell_address = { c: C, r: R };
        var data = XLSX.utils.encode_cell(cell_address);
        program.push(tab1[data]);
    }
    PdataRange.push(program);
}
var newprogramcsv = XLSX.utils.book_new();
var programcsv = XLSX.utils.aoa_to_sheet(PdataRange);

programcsv.A1.v = 'programAlias';
programcsv.A1.r = 'programAlias';
programcsv.A1.h = 'programAlias';
programcsv.A1.w = 'programAlias';
programcsv.B1.v = 'programName';
programcsv.B1.r = 'programName';
programcsv.B1.h = 'programName';
programcsv.B1.w = 'programName';
programcsv.C1.v = 'programCapacity';
programcsv.C1.r = 'programCapacity';
programcsv.C1.h = 'programCapacity';
programcsv.C1.w = 'programCapacity';
programcsv.D1.v = 'fundingEnabled';
programcsv.D1.r = 'fundingEnabled';
programcsv.D1.h = 'fundingEnabled';
programcsv.D1.w = 'fundingEnabled';
programcsv.E1.v = 'licenseNumber';
programcsv.E1.r = 'licenseNumber';
programcsv.E1.h = 'licenseNumber';
programcsv.E1.w = 'licenseNumber';


XLSX.utils.book_append_sheet(newprogramcsv, programcsv, "program");
XLSX.writeFile(newprogramcsv, "E:\\program.csv");



//classroom.csv
var CRrange = { s: { c: 1, r: 34 }, e: { c: 4, r: 39 } }; //B35:E40
var CRdataRange = [];
/* Iterate through each element in the structure */
for (var R = CRrange.s.r; R <= CRrange.e.r; ++R) {
    let classroom = [];
    for (var C = CRrange.s.c; C <= CRrange.e.c; ++C) {
        var cell_address = { c: C, r: R };
        var data = XLSX.utils.encode_cell(cell_address);
        classroom.push(tab1[data]);
    }
    CRdataRange.push(classroom);
}
var newclassroomcsv = XLSX.utils.book_new();
var classroomcsv = XLSX.utils.aoa_to_sheet(CRdataRange);

classroomcsv.A1.v = 'classroomName';
classroomcsv.A1.r = 'classroomName';
classroomcsv.A1.h = 'classroomName';
classroomcsv.A1.w = 'classroomName';
classroomcsv.B1.v = 'classroomCapacity';
classroomcsv.B1.r = 'classroomCapacity';
classroomcsv.B1.h = 'classroomCapacity';
classroomcsv.B1.w = 'classroomCapacity';
classroomcsv.C1.v = 'description';
classroomcsv.C1.r = 'description';
classroomcsv.C1.h = 'description';
classroomcsv.C1.w = 'description';


XLSX.utils.book_append_sheet(newclassroomcsv, classroomcsv, "classroom");
XLSX.writeFile(newclassroomcsv, "E:\\classroom.csv");



//fee.csv
var Frange = { s: { c: 1, r: 47 }, e: { c: 4, r: 53 } }; //B48:E54
var FdataRange = [];
/* Iterate through each element in the structure */
for (var R = Frange.s.r; R <= Frange.e.r; ++R) {
    let fee = [];
    for (var C = Frange.s.c; C <= Frange.e.c; ++C) {
        var cell_address = { c: C, r: R };
        var data = XLSX.utils.encode_cell(cell_address);
        fee.push(tab1[data]);
    }
    FdataRange.push(fee);
}
var newfeecsv = XLSX.utils.book_new();
var feecsv = XLSX.utils.aoa_to_sheet(FdataRange);

feecsv.A1.v = 'programName';
feecsv.A1.r = 'programName';
feecsv.A1.h = 'programName';
feecsv.A1.w = 'programName';
feecsv.B1.v = 'scheduleName';
feecsv.B1.r = 'scheduleName';
feecsv.B1.h = 'scheduleName';
feecsv.B1.w = 'scheduleName';
feecsv.C1.v = 'tuition';
feecsv.C1.r = 'tuition';
feecsv.C1.h = 'tuition';
feecsv.C1.w = 'tuition';

//console.log(feecsv);
XLSX.utils.book_append_sheet(newfeecsv, feecsv, "fee");
XLSX.writeFile(newfeecsv, "E:\\fee.csv");




//schedule
var Srange = { s: { c: 7, r: 21 }, e: { c: 27, r: 23 } }; //H22:AB24
var SdataRange = [];
/* Iterate through each element in the structure */
for (var R = Srange.s.r; R <= Srange.e.r; ++R) {
    let schedule = [];
    for (var C = Srange.s.c; C <= Srange.e.c; ++C) {
        var cell_address = { c: C, r: R };
        var data = XLSX.utils.encode_cell(cell_address);
        schedule.push(tab1[data]);
    }
    SdataRange.push(schedule) === 'N/A' ? '0' : schedule;
    SdataRange.push(schedule) === ' ' ? '0' : schedule;

    // if (SdataRange.cell == 'N/A') {
    //     SdataRange.cell.v === '0';
    //     return SdataRange;
    // }
}

var newschedulecsv = XLSX.utils.book_new();
var schedulecsv = XLSX.utils.aoa_to_sheet(SdataRange);

schedulecsv.A1.v = 'scheduleName';
schedulecsv.A1.r = 'scheduleName';
schedulecsv.A1.h = 'scheduleName';
schedulecsv.A1.w = 'scheduleName';

schedulecsv.B1.v = 'startTimeMonday';
schedulecsv.B1.r = 'startTimeMonday';
schedulecsv.B1.h = 'startTimeMonday';
schedulecsv.B1.w = 'startTimeMonday';

schedulecsv.C1.v = 'startMinuteMonday';
schedulecsv.C1.r = 'startMinuteMonday';
schedulecsv.C1.h = 'startMinuteMonday';
schedulecsv.C1.w = 'startMinuteMonday';

schedulecsv.D1.v = 'endTimeMonday';
schedulecsv.D1.r = 'endTimeMonday';
schedulecsv.D1.h = 'endTimeMonday';
schedulecsv.D1.w = 'endTimeMonday';

schedulecsv.E1.v = 'endMinuteMonday';
schedulecsv.E1.r = 'endMinuteMonday';
schedulecsv.E1.h = 'endMinuteMonday';
schedulecsv.E1.w = 'endMinuteMonday';

schedulecsv.F1.v = 'startTimeTuesday';
schedulecsv.F1.r = 'startTimeTuesday';
schedulecsv.F1.h = 'startTimeTuesday';
schedulecsv.F1.w = 'startTimeTuesday';

schedulecsv.G1.v = 'startMinuteTuesday';
schedulecsv.G1.r = 'startMinuteTuesday';
schedulecsv.G1.h = 'startMinuteTuesday';
schedulecsv.G1.w = 'startMinuteTuesday';

schedulecsv.H1.v = 'endTimeTuesday';
schedulecsv.H1.r = 'endTimeTuesday';
schedulecsv.H1.h = 'endTimeTuesday';
schedulecsv.H1.w = 'endTimeTuesday';

schedulecsv.I1.v = 'endMinuteTuesday';
schedulecsv.I1.r = 'endMinuteTuesday';
schedulecsv.I1.h = 'endMinuteTuesday';
schedulecsv.I1.w = 'endMinuteTuesday';

schedulecsv.J1.v = 'startTimeWednesday';
schedulecsv.J1.r = 'startTimeWednesday';
schedulecsv.J1.h = 'startTimeWednesday';
schedulecsv.J1.w = 'startTimeWednesday';

schedulecsv.K1.v = 'startMinuteWednesday';
schedulecsv.K1.r = 'startMinuteWednesday';
schedulecsv.K1.h = 'startMinuteWednesday';
schedulecsv.K1.w = 'startMinuteWednesday';

schedulecsv.L1.v = 'endTimeWednesday';
schedulecsv.L1.r = 'endTimeWednesday';
schedulecsv.L1.h = 'endTimeWednesday';
schedulecsv.L1.w = 'endTimeWednesday';

schedulecsv.M1.v = 'endMinuteWednesday';
schedulecsv.M1.r = 'endMinuteWednesday';
schedulecsv.M1.h = 'endMinuteWednesday';
schedulecsv.M1.w = 'endMinuteWednesday';

schedulecsv.N1.v = 'startTimeThursday';
schedulecsv.N1.r = 'startTimeThursday';
schedulecsv.N1.h = 'startTimeThursday';
schedulecsv.N1.w = 'startTimeThursday';

schedulecsv.O1.v = 'startMinuteThursday';
schedulecsv.O1.r = 'startMinuteThursday';
schedulecsv.O1.h = 'startMinuteThursday';
schedulecsv.O1.w = 'startMinuteThursday';

schedulecsv.P1.v = 'endTimeThursday';
schedulecsv.P1.r = 'endTimeThursday';
schedulecsv.P1.h = 'endTimeThursday';
schedulecsv.P1.w = 'endTimeThursday';

schedulecsv.Q1.v = 'endMinuteThursday';
schedulecsv.Q1.r = 'endMinuteThursday';
schedulecsv.Q1.h = 'endMinuteThursday';
schedulecsv.Q1.w = 'endMinuteThursday';

schedulecsv.R1.v = 'startTimeFriday';
schedulecsv.R1.r = 'startTimeFriday';
schedulecsv.R1.h = 'startTimeFriday';
schedulecsv.R1.w = 'startTimeFriday';

schedulecsv.S1.v = 'startMinuteFriday';
schedulecsv.S1.r = 'startMinuteFriday';
schedulecsv.S1.h = 'startMinuteFriday';
schedulecsv.S1.w = 'startMinuteFriday';

schedulecsv.T1.v = 'endTimeFriday';
schedulecsv.T1.r = 'endTimeFriday';
schedulecsv.T1.h = 'endTimeFriday';
schedulecsv.T1.w = 'endTimeFriday';

schedulecsv.U1.v = 'endMinuteFriday';
schedulecsv.U1.r = 'endMinuteFriday';
schedulecsv.U1.h = 'endMinuteFriday';
schedulecsv.U1.w = 'endMinuteFriday';

XLSX.utils.book_append_sheet(newschedulecsv, schedulecsv, "schedule");
XLSX.writeFile(newschedulecsv, "E:\\schedule.csv");


fs.readFile("E:\\schedule.csv", 'utf8', function(err, Sdata) {
    if (err) {
        return console.log(err);
    }
    var SFV = Sdata.replace('N/A', '0');

    fs.writeFile("E:\\schedule.csv", SFV, 'utf8', function(err) {
        if (err) return console.log(err);
    });
});



//child
//console.log(childdata);
var newchilddata = childdata.map(function(cr) {
    delete cr.id;
    cr.birthDay = cr.birthDay; //.toISOString.replace(/\T.+/, '').trim();
    cr.startDate = cr.startDate; //.replace(/\T.+/, '').trim();
    cr.Program = cr.Program.trim();
    cr.familyEmail = cr.familyEmail.trim();
    cr.firstName = cr.firstName.trim();
    cr.middleName = cr.middleName.trim();
    cr.lastName = cr.lastName.trim();
    cr.nickName = cr.nickName.trim();
    return cr;
});

//console.log(newchilddata);

var newchildcsv = XLSX.utils.book_new();
var newchildtab = XLSX.utils.json_to_sheet(newchilddata);
XLSX.utils.book_append_sheet(newchildcsv, newchildtab, "child");
XLSX.writeFile(newchildcsv, "E:\\child.csv");


//family
var newfamilydata = familydata.map(function(fr) {
    delete fr.id;
    fr.email = fr.email.toString().trim();
    fr.name = fr.name.trim();
    fr.email = fr.email.trim();
    fr.cellPhone = fr.cellPhone.toString().replace(/\D/g, "").trim();
    if (fr.cellPhone.length > 10) {
        fr.cellPhone = fr.cellPhone
    } else {
        fr.cellPhone = [fr.cellPhone.slice(0, 3), '-', fr.cellPhone.slice(3, 6), '-', fr.cellPhone.slice(6)].join('');
    }
    return fr;
});

//delete duplicate rows
var newfamilydata = newfamilydata.reduce((unique, o) => {
    if (!unique.some(fr => fr.name === o.name && fr.cellPhone === o.cellPhone)) {
        unique.push(o);
    }
    return unique;
}, []);


//console.log(newfamilydata);
var newfamilycsv = XLSX.utils.book_new();
var newfamilytab = XLSX.utils.json_to_sheet(newfamilydata);
XLSX.utils.book_append_sheet(newfamilycsv, newfamilytab, "family");
XLSX.writeFile(newfamilycsv, "E:\\family.csv");





//waiting
// var Wrange = waiting { s: { c: 1, r: 0 }, e: { c: 20, r: 2 } }; //B21 : F24
// var WdataRange = [];
// /* Iterate through each element in the structure */
// for (var R = Wrange.s.r; R <= Wrange.e.r; ++R) {
//     let wait = [];
//     for (var C = Wrange.s.c; C <= Wrange.e.c; ++C) {
//         var cell_address = { c: C, r: R };
//         var data = XLSX.utils.encode_cell(cell_address);
//         wait.push(tab1[data]);
//     }
//     WdataRange.push(wait);
// }
// var newwaitingcsv = XLSX.utils.book_new();
// var waitingcsv = XLSX.utils.aoa_to_sheet(WdataRange);
// XLSX.utils.book_append_sheet(newwaitingcsv, waitingcsv, "Waiting");
// XLSX.writeFile(newwaitingcsv, "E:\\waiting.csv");

let newwaitingdata = waitingdata.map(function(waitingr) {
    delete waitingr.id;
    waitingr.firstName = waitingr.firstName.trim();
    waitingr.middleName = waitingr.middleName.trim();
    waitingr.lastName = waitingr.lastName.trim();
    waitingr.nickName = waitingr.nickName.trim();
    waitingr.familyName = waitingr.familyName.trim();
    waitingr.familyEmail = waitingr.familyEmail.trim();
    waitingr.familyCell = waitingr.familyCell.toString().replace(/\D/g, "").trim();
    if (waitingr.familyCell.length > 10) {
        waitingr.familyCell = waitingr.familyCell
    } else {
        waitingr.familyCell = [waitingr.familyCell.slice(0, 3), '-', waitingr.familyCell.slice(3, 6), '-', waitingr.familyCell.slice(6)].join('');
    }
    return waitingr;
});
//console.log(newwaitingdata);

var newwaitingcsv = XLSX.utils.book_new();
var newwaitingtab = XLSX.utils.json_to_sheet(newwaitingdata);

XLSX.utils.book_append_sheet(newwaitingcsv, newwaitingtab, "Waiting");
XLSX.writeFile(newwaitingcsv, "E:\\waiting.csv");






//t1
var newt1data = t1data.map(function(t1r) {
    delete t1r.__EMPTY;
    t1r.birthDay = t1r.birthDay.toISOString().replace(/\T.+/, '');
    return t1r;
});
//console.log(newt1data);
var newt1csv = XLSX.utils.book_new();
var newt1tab = XLSX.utils.json_to_sheet(newt1data);
XLSX.utils.book_append_sheet(newt1csv, newt1tab, "Tuition-1");
XLSX.writeFile(newt1csv, "E:\\tuition1.csv");


//t2
var newt2data = t2data.map(function(t2r) {
    delete t2r.__EMPTY;
    t2r.birthDay = t2r.birthDay.toISOString().replace(/\T.+/, '');
    return t2r;
});
//console.log(newt1data);
var newt2csv = XLSX.utils.book_new();
var newt2tab = XLSX.utils.json_to_sheet(newt2data);
XLSX.utils.book_append_sheet(newt2csv, newt2tab, "Tuition-2");
XLSX.writeFile(newt2csv, "E:\\tuition2.csv");


//t3
var newt3data = t3data.map(function(t3r) {
    delete t3r.__EMPTY;
    t3r.birthDay = t3r.birthDay.toISOString().replace(/\T.+/, '');
    return t3r;
});
//console.log(newt1data);
var newt3csv = XLSX.utils.book_new();
var newt3tab = XLSX.utils.json_to_sheet(newt3data);
XLSX.utils.book_append_sheet(newt3csv, newt3tab, "Tuition-3");
XLSX.writeFile(newt3csv, "E:\\tuition3.csv");


//t4
var newt4data = t4data.map(function(t4r) {
    delete t4r.__EMPTY;
    t4r.birthDay = t4r.birthDay.toISOString().replace(/\T.+/, '');
    return t4r;
});
//console.log(newt1data);
var newt4csv = XLSX.utils.book_new();
var newt4tab = XLSX.utils.json_to_sheet(newt4data);
XLSX.utils.book_append_sheet(newt4csv, newt4tab, "Tuition-4");
XLSX.writeFile(newt4csv, "E:\\tuition4.csv");


//t5
var newt5data = t5data.map(function(t5r) {
    delete t5r.__EMPTY;
    t5r.birthDay = t5r.birthDay.toISOString().replace(/\T.+/, '');
    return t5r;
});
//console.log(newt1data);
var newt5csv = XLSX.utils.book_new();
var newt5tab = XLSX.utils.json_to_sheet(newt5data);
XLSX.utils.book_append_sheet(newt5csv, newt5tab, "Tuition-5");
XLSX.writeFile(newt5csv, "E:\\tuition5.csv");


//t6
var newt6data = t6data.map(function(t6r) {
    delete t6r.__EMPTY;
    t6r.birthDay = t6r.birthDay.toISOString().replace(/\T.+/, '');
    return t6r;
});
//console.log(newt1data);
var newt6csv = XLSX.utils.book_new();
var newt6tab = XLSX.utils.json_to_sheet(newt6data);
XLSX.utils.book_append_sheet(newt6csv, newt6tab, "Tuition-6");
XLSX.writeFile(newt6csv, "E:\\tuition6.csv");


//t7
var newt7data = t7data.map(function(t7r) {
    delete t7r.__EMPTY;
    t7r.birthDay = t7r.birthDay.toISOString().replace(/\T.+/, '');
    return t7r;
});
//console.log(newt1data);
var newt7csv = XLSX.utils.book_new();
var newt7tab = XLSX.utils.json_to_sheet(newt7data);
XLSX.utils.book_append_sheet(newt7csv, newt7tab, "Tuition-7");
XLSX.writeFile(newt7csv, "E:\\tuition7.csv");


//t8
var newt8data = t8data.map(function(t8r) {
    delete t8r.__EMPTY;
    t8r.birthDay = t8r.birthDay.toISOString().replace(/\T.+/, '');
    return t8r;
});
//console.log(newt1data);
var newt8csv = XLSX.utils.book_new();
var newt8tab = XLSX.utils.json_to_sheet(newt8data);
XLSX.utils.book_append_sheet(newt8csv, newt8tab, "Tuition-8");
XLSX.writeFile(newt8csv, "E:\\tuition8.csv");


//t9
var newt9data = t9data.map(function(t9r) {
    delete t9r.__EMPTY;
    t9r.birthDay = t9r.birthDay.toISOString().replace(/\T.+/, '');
    return t9r;
});
//console.log(newt1data);
var newt9csv = XLSX.utils.book_new();
var newt9tab = XLSX.utils.json_to_sheet(newt9data);
XLSX.utils.book_append_sheet(newt9csv, newt9tab, "Tuition-9");
XLSX.writeFile(newt9csv, "E:\\tuition9.csv");


//t10
var newt10data = t10data.map(function(t10r) {
    delete t10r.__EMPTY;
    t10r.birthDay = t10r.birthDay.toISOString().replace(/\T.+/, '');
    return t10r;
});
//console.log(newt1data);
var newt10csv = XLSX.utils.book_new();
var newt10tab = XLSX.utils.json_to_sheet(newt10data);
XLSX.utils.book_append_sheet(newt10csv, newt10tab, "Tuition-10");
XLSX.writeFile(newt10csv, "E:\\tuition10.csv");


//t11
var newt11data = t11data.map(function(t11r) {
    delete t11r.__EMPTY;
    t11r.birthDay = t11r.birthDay.toISOString().replace(/\T.+/, '');
    return t11r;
});
//console.log(newt1data);
var newt11csv = XLSX.utils.book_new();
var newt11tab = XLSX.utils.json_to_sheet(newt11data);
XLSX.utils.book_append_sheet(newt11csv, newt11tab, "Tuition-11");
XLSX.writeFile(newt11csv, "E:\\tuition11.csv");


//t12
var newt12data = t12data.map(function(t12r) {
    delete t12r.__EMPTY;
    t12r.birthDay = t12r.birthDay.toISOString().replace(/\T.+/, '');
    return t12r;
});
//console.log(newt12data);
var newt12csv = XLSX.utils.book_new();
var newt12tab = XLSX.utils.json_to_sheet(newt12data);
XLSX.utils.book_append_sheet(newt12csv, newt12tab, "Tuition-12");
XLSX.writeFile(newt12csv, "E:\\tuition12.csv");