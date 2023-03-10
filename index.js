import {numToString} from "./script.js";
const button = document.querySelector("#clickme")
const label = document.querySelector("#resultLabel")
const input = document.querySelector("#numberInput")

button.addEventListener("click", event=>{
    label.textContent = numToString(input.value)
})

