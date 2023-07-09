import mongoose from "mongoose";

const connectDB = async () => {
    try {
        await mongoose.connect(process.env.MONGODB_CONNECT_URI);
        console.log("Connected to MongoDB Succesfully");
    } catch (error) {
        console.log("Connect failed" + error);
    }
}

export default connectDB;