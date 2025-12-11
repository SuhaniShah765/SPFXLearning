import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Employee } from "../models/Employee";

export default class EmployeeService {

  private context: WebPartContext;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.context = context;
    this.listName = listName;
  }

  public async getAllEmployees(): Promise<Employee[]> {

    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=
      Id,Title,EmployeeCode,Department,JobTitle,Email,ProfilePhoto,
      PhoneNumber,Location,Manager/Title,JoiningDate,WorkAnniversary,
      Skills,Status,AboutMe&$expand=Manager`;

    console.log("Fetching:", endpoint);

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    const data = await response.json();
    console.log("SP LIST DATA:", data.value);

    return data.value.map((i: any) => ({
      Id: i.Id,
      Title: i.Title,
      EmployeeCode: i.EmployeeCode,
      Department: i.Department,
      JobTitle: i.JobTitle,
      Email: i.Email,
      Photo: i.ProfilePhoto,              // 🔥 FIXED field name
      PhoneNumber: i.PhoneNumber,
      Location: i.Location,
      Manager: i.Manager?.Title,          // 🔥 Expand person field
      JoiningDate: i.JoiningDate,
      WorkAnniversary: i.WorkAnniversary,
      Skills: i.Skills,
      Status: i.Status,
      AboutMe: i.AboutMe
    }));
  }
}
