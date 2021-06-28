const express= require("express")
const https = require("https")
const fs= require("fs");
const path =require("path")
var fetch = require("node-fetch")
var formurlencoded = require('form-urlencoded');

const cert = fs.readFileSync("C:/Users/prdalmia/.office-addin-dev-certs/localhost.crt");
const key =fs.readFileSync("C:/Users/prdalmia/.office-addin-dev-certs/localhost.key");
const app=express();

const server = https.createServer({key:key, cert:cert},app);

app.use(express.static(__dirname+'/public'));

app.get("/",(req,resp) => {resp.sendFile(__dirname+"index.html")});

server.listen("3001", ()=>{console.log("Listening on port 3001")});

