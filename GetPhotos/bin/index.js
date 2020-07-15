#!/usr/bin/env node

const yargs = require("yargs")
const axios = require("axios")
const qs = require("qs")
const fs = require('fs')
const path = require('path')

const options = yargs
    .scriptName("UserPhotos")
    .usage("Usage: --ci <client_id> --cs <client_secret> --ti <tenant_id>")
    .option("ci", {
        alias: "client_id",
        describe: "The client_id of the Azure API",
        type: "string",
        demandOption: true
    })
    .option("cs", {
        alias: "client_secret",
        describe: "The client_secret of the Azure API",
        type: "string",
        demandOption: true
    })
    .option("ti", {
        alias: "tenant_id",
        describe: "The tenant_id of the Azure API",
        type: "string",
        demandOption: true
    })
    .argv;

const postAppAccessTokenURL = `https://login.microsoftonline.com/${options.ti}/oauth2/v2.0/token`
const getUsersStandard = 'https://graph.microsoft.com/v1.0/users'

const postAppAccessTokenData = qs.stringify({
    grant_type: 'client_credentials',
    client_id: options.ci,
    client_secret: options.cs,
    scope: 'https://graph.microsoft.com/.default'
})

const postAppAccessTokenHeaders = {
    'Content-Type': 'application/x-www-form-urlencoded'
}

async function postAppAccessToken() {
    try {
        console.log('Attempting to generate an Access Token.')
        const response = await axios.post(
            postAppAccessTokenURL,
            postAppAccessTokenData,
            postAppAccessTokenHeaders
        )
        const token = await response.data.access_token
        console.log('Access Token successfully generated.')
        return token
    } catch (error) {
        console.log('Error occured during generation of Access Token.')
        console.log(error)
    }
}

async function getUsers(token) {
    var allResults = []
    let url = 'https://graph.microsoft.com/v1.0/users'
    do {
        const response = await axios.get(url, {
            headers: {
                'Authorization': token
            }
        })
        const data = await response.data.value
        url = response.data['@odata.nextLink']
        allResults.push(...data)
    } while (url)
    return allResults
}

async function getUserPhoto(userPrincipalName, id, token) {
    const url = `https://graph.microsoft.com/v1.0/users/${id}/photo/$value`
    //Whatever folder the script is ran from will be where the images are saved.
    const pathlocation = path.resolve(`${userPrincipalName}.jpeg`)
    const writer = fs.createWriteStream(pathlocation)

    try {
        console.log(`Attempting to gather ${userPrincipalName}'s photo.`)

        const response = await axios({
            url,
            method: 'GET',
            responseType: 'stream',
            headers: {
                'Authorization': token
            }
        })

        response.data.pipe(writer)

        return new Promise((resolve, reject) => {
            writer.on('finish', resolve)
            writer.on('error', reject)
            console.log('Photo Collected.')
        })
        
    } catch (error) {
        if (error.response.status === 404) {
            console.log(`${userPrincipalName} does not have a photo.`)
            fs.unlink(pathlocation, (err) => {
                if (err) {
                    console.log(err)
                    return
                }
            })
        } else {
            console.log('Unknown error. Photo could not be collected.')
        }
    }
}

async function data() {
    try {

        const token = await postAppAccessToken()

        const users = await getUsers(token)

        for (const user of users) {
            if (user.givenName !== 'null' && user.surname !== 'null') {
                const getPhoto = await getUserPhoto(user.userPrincipalName, user.id, token)
            }
        }

    } catch (error) {
        console.log('Something broke')
        console.log(error)
    }
}

data()