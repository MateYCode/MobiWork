const XLSX = require('xlsx');
const axios = require('axios');
var eventService = require("./eventService.js");
var parseoXML = require('xml2js').parseString;
const express = require('express')
const path = require('path')
const fileUpload = require('express-fileupload')
const basicAuth = require("basic-auth");

const auth = (req, res, next) => {
      const unauthorized = (res) => {
            res.set("WWW-Authenticate", "Basic realm=Authorization Required");
            return res.sendStatus(401);
      };
      const user = basicAuth(req);
      if (!user || !user.name || !user.pass) {
            return unauthorized(res);
      };
      if (user.name === "user" && user.pass === "Password") {
            return next();
      } else {
            return unauthorized(res);
      };
};

const getDateFromExcel = excelDate => {
      /* Excel date to string */
      let jsDate = new Date((excelDate - (25567 + 1)) * 86400 * 1000);
      let mobiDate = jsDate.getFullYear() + '/' + (jsDate.getMonth() + 1) + '/' + jsDate.getDate()
      return mobiDate
}

const getTimeFromExcel = excelTime => {
      /* Excel time to string */
      let hour = Math.floor(excelTime * 24).toString();
      if (hour.length < 2) {
            hour = '0' + hour;
      }
      let minute = Math.round(((excelTime * 24) % 1) * 60).toString();
      if (minute.length < 2) {
            minute = '0' + minute;
      }
      return mobiTime = (hour + ':' + minute);
}

const app = express()
app.set('view engine', 'ejs')
app.set('views', path.join(__dirname, 'views'))
app.get('/', auth, express.static(__dirname + '/views'));
app.use(fileUpload())
app.post('/upload', (req, res) => {
      let archivo;
      let errors = [];
      let newWO = 0;
      archivo = req.files.file
      archivo.mv(`./files/${archivo.name}`, (err) => {
            if (err) return res.status(500).send({ message: err })
            const wb = XLSX.readFile('./files/' + archivo.name);
            let hoja = wb.Props.SheetNames[0]
            const ws = wb.Sheets[hoja];
            const impoExcelJson = XLSX.utils.sheet_to_json(ws, { raw: true });
            let woTypes = {
                  WALMART: 3658,
                  EJEMPLO: 9317,
                  OT1: 9917,
                  OT2: 9918,
                  OT3: 9919,
                  OT5: 9920
            }
            let token;
            let clients = [];
            eventService.endPoints.authenticate()
                  .then(response => {
                        parseoXML(response.data, (err, result) => {
                              token = result.response.token[0];
                        })
                        return token
                  })
                  .then(token => {
                        return eventService.endPoints.clientList(token)
                  })
                  .then(resp => {
                        let arrayClient;
                        parseoXML(resp.data, (err, result) => {
                              arrayClient = result.response.customerList[0].customer;
                              if (err) {
                                    errors.push(err)
                              }
                        })
                        return arrayClient
                  })
                  .then(arrayClient => {
                        arrayClient.forEach(item => {
                              let client = []; let idsClientAndLocation = []
                              client.push(item.companyName[0])
                              idsClientAndLocation.push(item.mobiworkCustomerId[0])
                              idsClientAndLocation.push(item.address[0].addressId[0])
                              client.push(idsClientAndLocation) //client= [client,[id client,id ubicacion]]
                              clients.push(client)
                        })
                        return clients
                  })
                  .then(() => {
                        return eventService.endPoints.userList(token)
                  })
                  .then((resp) => {
                        let users; let userId; let resultado;
                        parseoXML(resp.data, (err, result) => {
                              resultado = result.response.userList[0].user;
                              if (err) {
                                    errors.push(err)
                              }
                        })
                        return resultado;
                  })
                  .then(users => {
                        let clientsMap = new Map(clients)
                        let promises = impoExcelJson.map((fila, index) => {
                              let userId;
                              let userName = fila.Nombre;
                              let customer = fila.Empresa
                              let idUbicacion; let idCliente; let mobiDate; let mobiTime; let woType
                              function userExists() {
                                    if (!userId && userName) {
                                          //User not found in the database
                                          errors.push(`Fila: ${index + 2} .Usuario: ${userName} no coincide con un usuario en MobiWork. La orden no es creada.`);
                                          return false;
                                    }
                                    else {
                                          return true
                                    }
                              }
                              function customerExists() {
                                    if (clientsMap.get(customer) != null) { //customer not found in the database
                                          idCliente = clientsMap.get(customer)[0]
                                          idUbicacion = clientsMap.get(customer)[1]
                                          mobiDate = getDateFromExcel(fila['Fecha (DD/MM/AAAA)'])
                                          mobiTime = getTimeFromExcel(fila['Hora (hh:mm)']);
                                          woType = woTypes[fila.Tipo];
                                          users.forEach((item) => {
                                                let userNameMobi = item.firstName[0] + ' ' + item.lastName[0];
                                                if (!userId) {
                                                      if (userName == userNameMobi) {
                                                            userId = item.mobiworkUserId[0]
                                                      }
                                                }
                                          })
                                          return true
                                    }
                                    else {
                                          errors.push(`Fila: ${index + 2} .Cliente: ${customer} no coincide con un client en MobiWork. La orden no es creada`)
                                          return false;
                                    }
                              }

                              if (customerExists() && userExists()) {
                                    let request = `<request><workOrder><mobiworkWorkOrderId></mobiworkWorkOrderId><externalId></externalId><customerId>${(idCliente)}</customerId><customerName>${(customer)}</customerName><description>${(fila.Descripcion)}</description><workOrderTypeId>${woType}</workOrderTypeId><workOrderType></workOrderType><status>ASSIGNED</status><customStatusId></customStatusId><customStatus></customStatus><location><addressId>${(idUbicacion)}</addressId><address1>Direccion 123</address1><address2></address2><city>${fila.Ciudad}</city><state>${fila.Provincia}</state><zipCode></zipCode><countryId>4</countryId><latitude></latitude><longitude></longitude><name></name><baseDistance></baseDistance></location><createdDate></createdDate><customFields></customFields><scheduleList><schedule><userId>${(userId)}</userId><date>${(mobiDate)}</date><time>${mobiTime}</time><duration>30</duration></schedule></scheduleList></workOrder></request>`
                                    return eventService.endPoints.woAdd(token, request) //to create the work order
                                          .then(rta => {
                                                parseoXML(rta.data, (err, result) => {
                                                      if (result.response.$.statusCode != 1) {
                                                            errors.push(JSON.stringify(result.response.error[0]._))
                                                      }
                                                })
                                          })
                                          .catch(err => {
                                                errors.push(err)
                                          })
                              }
                        }) //map for each Excel row   
                        return Promise.all(promises)
                  })
                  .then(results => {
                        console.log('Listo errores:', errors);
                        res.render('errors', { created: newWO, errorList: errors });
                  })
                  .catch(err => { console.log('Errores', err.data) })

      })



})


app.listen(3000, () => console.log('Running'))