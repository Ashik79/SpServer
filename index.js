const express = require('express')
const app = express()
const port = process.env.PORT || 5000;
const cors = require("cors");

const XLSX = require('xlsx');


app.use(express.json({ limit: '10mb' })); // Adjust the limit as needed
app.use(express.urlencoded({ limit: '10mb', extended: true }));
app.use(cors());

app.listen(port, () => {
  // console.log("port is", port)
})
app.get('/', (req, res) => {
  res.send("server isff running")
})
const { MongoClient, ServerApiVersion, ObjectId } = require('mongodb');

const uri = `mongodb+srv://brandshop:brandshop1212@cluster0.ak91fsl.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0`
// Create a MongoClient with a MongoClientOptions object to set the Stable API version
const client = new MongoClient(uri, {
  serverApi: {
    version: ServerApiVersion.v1,
    strict: true,
    deprecationErrors: true,
  }
});

const database = client.db("spoffice");
const studentsCollection = database.collection("students");
studentsCollection.createIndex({ id: 1 }, { unique: true });
const usersCOllection = database.collection("users")
const examsCollection = database.collection("exams")
const couponCollection = database.collection("coupons")
const courseCollection = database.collection("courses")


async function run() {
  try {

    //get all students
    app.get('/students', async (req, res) => {

      const cursor = studentsCollection.find();
      const students = await cursor.toArray();
      res.send(students);
    })


    //get a student
    app.get('/student/:id', async (req, res) => {
      const id = req.params.id
      const query = {
        id: id
      }
      const result = await studentsCollection.findOne(query)
      if (result) {
        res.send(result);
      } else {
        res.status(404).send({ message: 'Student not found' });
      }
    })
    //get a staff
    app.get('/staff/:id', async (req, res) => {
      const id = req.params.id
      const query = {
        _id: new ObjectId(id)
      }
      const result = await usersCOllection.findOne(query)
      if (result) {
        res.send(result);
      } else {
        res.status(404).send({ message: 'Student not found' });
      }
    })
    // students er number er array paite ids er array diye
    app.post('/getnumbers', async (req, res) => {
      const ids = req.body
      const students = await studentsCollection
        .find({ id: { $in: ids } },)
        .toArray();

      const phoneNumbers = []
      students.map(student => {

        const obj = {
          phone: student.phone,
          name: student.name
        }
        phoneNumbers.push(obj)
      });
      // console.log(phoneNumbers)

      // send the phone numbers and names
      JSON.stringify(phoneNumbers)
      res.send(phoneNumbers);
    })
    // students er  array paite ids er array diye
    app.post('/getstudents', async (req, res) => {
      const ids = req.body.ids
      const students = await studentsCollection
        .find({ id: { $in: ids } },)
        .toArray();



      res.send(students);
    })


    // Search students 
    app.post('/students', async (req, res) => {
      const query = req.body
      const cursor = studentsCollection.find(query)
      const result = await cursor.toArray()
      res.send(result)
    })
    // student overview 
    app.post('/studentoverview', async (req, res) => {
      const date = req.body.date

      const cursor = studentsCollection.find()
      const result = await cursor.toArray()

      const admitted = result.filter(student => student.programs.length != 0)
      const registeredToday = result.filter(student => student.admissionDate == date)
      const admittedToday = []
      admitted?.map(student => {
        const temp = student.programs?.some(program => program.payDate == date)
        if (temp) {
          admittedToday.push(student)
        }
      })
      const data = {
        total: result.length,
        admitted: admitted.length,
        registeredToday: registeredToday.length,
        admittedToday: admittedToday.length
      }

      res.send(data)
    })
    // Payment OVerview download
    app.post('/download/overview', async (req, res) => {
      const array = req.body
      const payments = array.map(payment => ({
        Id: payment.id,
        Name: payment.name,
        Amount: payment.pamount,
        Type: payment.type,
        Program: `${payment.program ? payment.program : ""}`,
        Month: payment.pmonth,
        Year: payment.pyear,
        Taken: payment.payDate,
        Taken_by: payment.ptaken,

      }))
      // console.log(payments)
      const worksheet = XLSX.utils.json_to_sheet(payments);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'payments');

      const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
      const fileName = `payments.xlsx`;

      res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
      res.setHeader('Content-Type', 'application/octet-stream');
      res.send(buffer);

    })

    //Student info Download
    app.post('/download/students', async (req, res) => {
      const array = req.body
      const students = array.map(student => ({
        Id: student.id,
        Name: student.name,
        Phone: student.phone,


      }))

      const worksheet = XLSX.utils.json_to_sheet(students);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'students');

      const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
      const fileName = `students.xlsx`;

      res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
      res.setHeader('Content-Type', 'application/octet-stream');
      res.send(buffer);

    })

    

    //Exam er all results Download
    const XLSX = require('xlsx');

    app.post('/download/examresults', async (req, res) => {
      const array = req.body;

      const results = array.map(result => ({
        Id: result.id,
        Name: result.name,
        MCQ: result.mcqMarks,
        Writen: result.writenMarks,
        Total: result.total,
        Merit: result.merit
      }));

      // Calculate total marks for the exam (same for all results)
      const totalMcq = array.length > 0 ? array[0].mcqTotal : 0;
      const totalWriten = array.length > 0 ? array[0].writenTotal : 0;
      const totalResults = array.length > 0 ? array.length : 0;
      const totalMarks = totalMcq + totalWriten;

      const worksheet = XLSX.utils.json_to_sheet(results);

      // Modify the headers
      worksheet['A1'].v = 'Id';
      worksheet['B1'].v = 'Name';
      worksheet['C1'].v = `MCQ (out of ${totalMcq})`;
      worksheet['D1'].v = `Writen (out of ${totalWriten})`;
      worksheet['E1'].v = `Total (out of ${totalMarks})`;
      worksheet['F1'].v = `Merit (out of ${totalResults})`;

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'results');

      const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
      const fileName = `results.xlsx`;

      res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
      res.setHeader('Content-Type', 'application/octet-stream');
      res.send(buffer);
    });

    //Student Result Download
    app.post('/download/results', async (req, res) => {
      const array = req.body
      const students = array.map(student => ({
        Title: student.title,
        Date: student.date,
        MCQ: `${student.mcqMarks} / ${student.mcqTotal}`,
        Writen: `${student.writenMarks} / ${student.writenTotal}`,
        Total: `${student.writenMarks + student.mcqMarks} / ${student.writenTotal + student.mcqTotal}`,



      }))

      const worksheet = XLSX.utils.json_to_sheet(students);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'students');

      const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
      const fileName = `results.xlsx`;

      res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);
      res.setHeader('Content-Type', 'application/octet-stream');
      res.send(buffer);

    })

    //ekta exam paite
    app.get('/getexam/:id', async (req, res) => {
      const id = req.params.id
      // console.log(id)
      const query = { _id: new ObjectId(id) }
      const result = await examsCollection.findOne(query)
      res.send(result)
    })
    //ekta video course paite
    app.get('/getvideocourse/:id', async (req, res) => {
      const id = req.params.id
      // console.log(id)
      const query = { _id: new ObjectId(id) }
      const result = await courseCollection.findOne(query)
      res.send(result)
    })
    //ekta Coupon paite
    app.get('/getcoupon/:code', async (req, res) => {
      const code = req.params.code

      const query = { code: code }
      const result = await couponCollection.findOne(query)
      if (result == null) {
        res.send({})
      }
      else res.send(result)
    })


    //exam page a sob exam load kora
    app.get('/getexams', async (req, res) => {

      const cursor = examsCollection.find()
      const result = await cursor.toArray()
      res.send(result)
    })
    //Courses page a sob course load kora
    app.get('/getcourses', async (req, res) => {

      const cursor = courseCollection.find()
      const result = await cursor.toArray()
      res.send(result)
    })
    app.get('/getcoupons', async (req, res) => {

      const cursor = couponCollection.find()
      const result = await cursor.toArray()
      res.send(result)
    })

    //exam add korar post
    app.post('/addexam', async (req, res) => {
      const exam = req.body;
      const result = await examsCollection.insertOne(exam)
      res.send(result)
    })

    app.post('/addcoupon', async (req, res) => {
      const coupon = req.body;
      const result = await couponCollection.insertOne(coupon)
      res.send(result)
    })

    //Course add korar post
    app.post('/addcourse', async (req, res) => {
      const course = req.body;
      const result = await courseCollection.insertOne(course)
      res.send(result)
    })
  

    //exam a result add korar post
    app.post('/exam/addresult/:id', async (req, res) => {
      const examId = req.params.id
      const data = req.body;
      const query = { _id: new ObjectId(examId) }
      const exam = await examsCollection.findOne(query)
      const duplicate = exam.results.some(result => result.id == data.id);
      if (duplicate) {
        return res.status(400).send('Duplicate result for the same student ');
      }

      exam.results.push(data);

      const updateResult = await examsCollection.updateOne(query, { $set: { results: exam.results } });
      if (updateResult.modifiedCount === 1) {
        res.status(200).send('Result added successfully');
      } else {
        res.status(500).send('Failed to add result');
      }



      // console.log(data, examId, exam,)
    })

    //user add korar post
    app.post('/adduser', async (req, res) => {
      const user = req.body;
      const result = await usersCOllection.insertOne(user)
      res.send(result)
    })

    //Update korte
    app.put('/student/update/:id', async (req, res) => {
      const id = req.params.id;
      const updatedData = {
        $set: req.body
      }
      const filter = {
        id: id
      }
      const result = await studentsCollection.updateOne(filter, updatedData)
      res.send(result)

    })
    //Exam result Update korte
    app.put('/exam/update/:id', async (req, res) => {
      const id = req.params.id;
      const updatedData = {
        $set: req.body
      }
      const filter = {
        _id: new ObjectId(id)
      }
      const result = await examsCollection.updateOne(filter, updatedData)
      res.send(result)

    })

    //eksathe students profile a results add korte

    app.put('/addbulkresults', async (req, res) => {
      const updates = req.body; // This will be an array of student updates

      // Create an array of bulk operations
      const bulkOps = updates.map(update => ({
        updateOne: {
          filter: { id: update.id },
          update: { $push: { exams: { $each: update.exams } } }
        }
      }));

      try {
        // Execute bulk operations
        const result = await studentsCollection.bulkWrite(bulkOps);
        res.send(result);
      } catch (error) {
        console.error('Error updating students:', error);
        res.status(500).send('Error updating students');
      }
    });


    //exam delete korte

    app.delete('/exam/delete/:id', async (req, res) => {
      const id = req.params.id
      const query = { _id: new ObjectId(id) }
      const result = await examsCollection.deleteOne(query)
      res.send(result)
    })
    //course delete korte

    app.delete('/course/delete/:id', async (req, res) => {
      const id = req.params.id
      const query = { _id: new ObjectId(id) }
      const result = await courseCollection.deleteOne(query)
      res.send(result)
    })
    app.delete('/coupon/delete/:id', async (req, res) => {
      const id = req.params.id
      const query = { _id: new ObjectId(id) }
      const result = await couponCollection.deleteOne(query)
      res.send(result)
    })
    //student delete korte

    app.delete('/student/delete/:id', async (req, res) => {
      const id = req.params.id
      const query = { id: id }
      const result = await studentsCollection.deleteOne(query)
      res.send(result)
    })

    // User er name nite
    app.get('/getuser/:param', async (req, res) => {
      const mail = req.params.param;
      if (mail != "null") {
        // console.log(mail)
        const query = {
          email: mail
        }
        const user = await usersCOllection.findOne(query)
        res.send(user)

      }

    })

    //admission post
    app.post('/admit', async (req, res) => {
      const admissionData = req.body;

      // console.log(admissionData)
      try {
        const result = await studentsCollection.insertOne(admissionData);
        res.send(result);
      } catch (error) {
        // Handle duplicate key error (11000) explicitly
        if (error.code === 11000) {
          res.status(405).send("Duplicate key error: ID must be unique.");
        } else {
          res.status(500).send("Error inserting data into MongoDB.");
        }
      }
    })
    //sob user er name array
    app.get("/getusers", async (req, res) => {
      const cursor = usersCOllection.find()
      const allUsers = await cursor.toArray()
      var allNames = []
      allUsers.map(user => {
        allNames.push(user.name)
      })
      const respond = JSON.stringify(allNames)
      res.send(respond)
    })
    //sob user er  array
    app.get("/getusersfull", async (req, res) => {
      const cursor = usersCOllection.find()
      const allUsers = await cursor.toArray()

      const respond = JSON.stringify(allUsers)
      res.send(respond)
    })

    //payment OVerview pete
    app.post('/api/payments', async (req, res) => {
      const { month, day, year, taker } = req.body;

      // console.log(taker)
      const cursor = studentsCollection.find()
      const allStudents = await cursor.toArray()
      var total = 0;
      var monthly = 0;
      var other = 0;
      var note = 0;
      var exam = 0;
      var course = 0;
      var admissionCount = 0;
      var monthlyCount = 0;
      const paymentArray = []

      allStudents.map(student => {
        if (student.payments && student.payments.length > 0) {
          student.payments.map(payment => {
            var dayMatch = false;
            var monthMatch = false;
            var yearMatch = false;
            var takerMatch = false;

            if (day) {
              if (payment.date == day)
                dayMatch = true
            }
            else if (!day) {
              dayMatch = true;
            }


            if (month) {
              if (payment.month == month) {
                monthMatch = true;
              }
            }
            else if (!month) {
              monthMatch = true;
            }

            if (year) {
              if (payment.year == year) {
                yearMatch = true;
              }
            }
            else if (!year) {
              yearMatch = true;
            }

            if (taker) {
              if (payment.ptaken == taker) {
                takerMatch = true;
              }
            }
            else if (!taker) {
              takerMatch = true;
            }



            if (dayMatch && monthMatch && takerMatch && yearMatch) {
              paymentArray.push(payment)
              total = total + parseInt(payment.pamount)
              if (payment.type == "Monthly") {
                monthly = monthly + parseInt(payment.pamount)
                monthlyCount++;
              }
              else if (payment.type == "Exam Fee" || payment.type == "Note Fee") {

                if (payment.type == "Exam Fee") {
                  admissionCount++;
                  exam += parseInt(payment.pamount)
                }
                else if (payment.type == 'Note Fee') {
                  note += parseInt(payment.pamount)
                }

              }
              else if (payment.type != "Monthly" && payment.type != "Exam Fee" && payment.type != "Note Fee") {
                other += parseInt(payment.pamount)
              }
            }

          })
        }
      })

      const data = { total, monthly, monthlyCount, note, exam, admissionCount, other, paymentArray }
      // console.log(data)
      const respond = JSON.stringify(data)
      res.send(respond)
    });

    //attendance add

    app.post('/attendance/:id', async (req, res) => {
      const id = req.params.id
      const attendance = req.body


      const student = await studentsCollection.findOne({ id: id })
      // console.log(student.attendances)

      const updatedAttendances = [...student.attendances, attendance]


      const result = await studentsCollection.updateOne({ id: id }, { $set: { attendances: updatedAttendances } })

      res.send(result)
    })

    //payment add
    app.put('/addpayment/:id', async (req, res) => {
      const data = req.body;
      const id = req.params.id;
      const filter = {
        id: id
      }
      const options = { upsert: true }
      if (data._id) {
        delete data._id;
      }
      const result = await studentsCollection.replaceOne(filter, data, options)
      res.send(result)



    })
    //staff update
    app.put('/updatestaff/:id', async (req, res) => {
      const data = req.body;
      // console.log(data)
      const id = req.params.id;
      const filter = {
        _id: new ObjectId(id)
      }
      const options = { upsert: true }
      if (data._id) {
        delete data._id;
      }
      const result = await usersCOllection.replaceOne(filter, data, options)
      res.send(result)



    })
    //course update
    app.put('/courseupdate/:id', async (req, res) => {
      const data = req.body;
      // console.log(data)
      const id = req.params.id;
      const filter = {
        _id: new ObjectId(id)
      }
      const options = { upsert: true }
      if (data._id) {
        delete data._id;
      }
      const result = await courseCollection.replaceOne(filter, data, options)
      res.send(result)



    })

    //exam er result delete
    app.put('/deleteresult/:id', async (req, res) => {
      const data = req.body;
      const id = req.params.id;
      const filter = {
        _id: new ObjectId(id)
      }
      // console.log(data)
      const options = { upsert: true }
      if (data._id) {
        delete data._id;
      }
      const result = await examsCollection.replaceOne(filter, data, options)
      res.send(result)



    })
    //Result add
    app.put('/addresult/:id', async (req, res) => {
      const data = req.body;
      const id = req.params.id;
      const filter = {
        id: id
      }
      const options = { upsert: true }
      if (data._id) {
        delete data._id;
      }
      const result = await studentsCollection.replaceOne(filter, data, options)
      res.send(result)

    })


  } finally {
    // Ensures that the client will close when you finish/error

  }
}



run().catch(console.dir);

