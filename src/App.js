import { AppBar, Container, Typography, Card, TextField, Button, Box, Backdrop, CircularProgress } from '@mui/material';
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import swal from 'sweetalert';

export default function App() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [subject, setSubject] = useState('');
  const [bodyHTML, setBodyHTML] = useState('');
  const [fileName, setFileName] = useState('Header of email column must be Email and name column must be Name in excel sheet');
  const [jsonData, setJsonData] = useState('');
  const [backdropOpen, setBackdropOpen] = useState(false);
  const [disabled, setDisabled] = useState(true);
  const [sendResponse, setSendResponse] = useState('Please wait');

  const sandMail = async (e) => {
    e.preventDefault();
    if (disabled) {
      swal({
        title: "Excel sheet not found",
        text: "Please upload excel sheet",
        icon: "info",
      })
    }
    else {
      setBackdropOpen(true);
      try {
        for (var i = 0; i < jsonData.length;) {
          const receiver = jsonData[i];
          const bodyHTMLWithNameReplaced = bodyHTML.replaceAll("{Name}", receiver.Name);
          const bodyHTMLWithBreakTagAdded = bodyHTMLWithNameReplaced.replaceAll("\n", "<br>");
          const formData = {
            senderEmail: email,
            appPassword: password,
            subject: subject,
            bodyHTML: bodyHTMLWithBreakTagAdded,
            receiverName: receiver.Name,
            receiverEmail: receiver.Email,
          }
          const requestOptions = {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify(formData)
          }
          const url = `https://mailhero.azurewebsites.net/api/sendMail`;
          await fetch(url, requestOptions)
            .then((response) => response.json())
            .then((data) => {
              if (data.status === 200) {
                setSendResponse(i + 1 + " : " + data.message);
                i++;
              }
              else {
                swal({
                  title: "Some error occured",
                  text: data.message,
                  icon: "error",
                }).then(() => {
                  setBackdropOpen(false);
                  setSendResponse('Please wait');
                  return
                });
              }
            })
        }
        swal({
          title: "All mail sent successfully",
          icon: "success",
        }).then(() => {
          setBackdropOpen(false)
          setSendResponse('Please wait');
        })
      }
      catch (error) {
        swal({
          title: "Some error occured",
          text: error,
          icon: "error",
        }).then(() => {
          setBackdropOpen(false);
          setSendResponse('Please wait');
          return
        });
      }
    }
  }

  const handleUpload = (e) => {
    e.preventDefault();
    var file = e.target.files[0];
    e.target.value = null;
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = e.target.result;
      let readedData = XLSX.read(data, { type: 'binary' });
      const wsname = readedData.SheetNames[0];
      const ws = readedData.Sheets[wsname];
      const dataParse = XLSX.utils.sheet_to_json(ws, { header: 0 });
      if (dataParse.length === 0 || dataParse[0].Email === undefined) {
        swal({
          title: "Email column not found",
          text: "Header of column containing email must be Email.",
          icon: "info",
        })
        setDisabled(true);
        setFileName('Header of email column must be Email and name column must be Name in excel sheet');
        setJsonData('');
      }
      else {
        setJsonData(dataParse);
        setFileName(file.name);
        setDisabled(false);
      }
    };
    reader.readAsBinaryString(file);
  }

  return (
    <>
      <Backdrop
        sx={{ color: '#fff', zIndex: (theme) => theme.zIndex.drawer + 1 }}
        open={backdropOpen}
      >
        <Box>
          <Box sx={{ justifyContent: "center", display: "flex" }}>
            <CircularProgress color="inherit" />
          </Box>
          <Typography sx={{ mt: 2 }}>{sendResponse}</Typography>
        </Box>
      </Backdrop>
      <AppBar position='static' sx={{ p: 2 }}>
        <Typography variant="h6">
          MailHero
        </Typography>
      </AppBar>
      <Container maxWidth="md" sx={{ py: 2 }}>
        <form onSubmit={sandMail}>
          <Card sx={{ p: 2 }}>
            <TextField label="Your gamil address" type="email" placeholder='xyz@gmail.com' fullWidth required sx={{ mb: 2 }} value={email} onChange={(e) => setEmail(e.target.value)} />
            <TextField label="App password" placeholder='Google Account >> Security >> App Passwords >> Generate' fullWidth required sx={{ mb: 2 }} value={password} onChange={(e) => setPassword(e.target.value)} />
            <TextField label="Subject of email" placeholder='Be clear and specific about the topic of the email' fullWidth required sx={{ mb: 2 }} value={subject} onChange={(e) => setSubject(e.target.value)} />
            <TextField label="Body of email" placeholder='In place of name of receiver write {Name} . Example: Dear Ram, => Dear {Name},' fullWidth required multiline minRows={8} sx={{ mb: 2 }} value={bodyHTML} onChange={(e) => setBodyHTML(e.target.value)} />
            <Box sx={{ display: 'flex', alignItems: 'center', mb: 2 }}>
              <TextField label="Excel file with email of receivers" value={fileName} fullWidth disabled={disabled} />
              <Button variant="outlined" component="label" sx={{ ml: 1, height: 54, textAlign: "center", width: 150 }}>
                Upload Excel
                <input type="file" onChange={(e) => handleUpload(e)} accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden />
              </Button>
            </Box>
            <Button variant="contained" type="submit" size="large">Send mail</Button>
          </Card>
        </form>
      </Container>
    </>
  )
}
