package org.example;

import jakarta.json.Json;
import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.GET;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import org.json.JSONObject;

/**
 * Root resource (exposed at "myresource" path)
 */
@Path("myresource")
public class MyResource {

    /**
     * Method handling HTTP GET requests. The returned object will be sent
     * to the client as "text/plain" media type.
     *
     * @return String that will be returned as a text/plain response.
     */
    @GET
    @Produces(MediaType.TEXT_PLAIN)
    public String getIt() {
        return "Got it!";
    }

    @POST
    @Produces(MediaType.APPLICATION_JSON)
    @Consumes(MediaType.APPLICATION_JSON)
    @Path("generateExcel")
    public Response generateExcel(String userData) {
        JSONObject json = new JSONObject(userData);
        String email = json.getString("email");
        String phone = json.getString("phone");
        String firstName = json.getString("firstName");
        String lastName = json.getString("lastName");

        User user = new User(email, phone, firstName, lastName);
        GenerateExcel generateExcel = new GenerateExcel();
        String filePath = generateExcel.addUserToExcel(user);
        JSONObject responseJson = new JSONObject();
        responseJson.put("filePath", filePath);
        return Response.ok(responseJson.toString(), MediaType.APPLICATION_JSON).build();
    }
}
