import express from 'express';
import cors from 'cors';

const app = express();

app.use(cors());
app.use(express.json());         

app.get('/',(req,res)=>{
    res.send('hello');
});

app.listen(3000, ()=>{
    console.log('server running onport 3000');
});
   