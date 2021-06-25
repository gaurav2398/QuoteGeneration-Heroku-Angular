import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Accounts } from '../Model/Accounts';
import { Questions } from '../Model/Questions';
import { QuestionsNew } from '../Model/QuestionsNew';
import { User } from '../Model/User';

@Injectable({
  providedIn: 'root'
})
export class UserService {
  status: string;
  constructor(private http: HttpClient) {}
  baseUrl: string = 'http://localhost:3000/'; // Base Url
  baseUrl1: string = 'https://quotegeneration-springboot.herokuapp.com/api/'; // Base Url
  baseUrl2: string = 'http://localhost:9091/api/'; // Base Url

  
  getAccounts() {
    return this.http.get<Accounts[]>(this.baseUrl1+'accounts');
  } // Get All accounts
  
  getUserById(createdBy: string) {
    return this.http.get<User>(this.baseUrl1 + 'accounts/' + createdBy);
  } // Get User By  Id

  createAccount(user: Accounts) {
    return this.http.post(this.baseUrl1+'accounts', user);
  } // Create New Insured Account

  getQuestions() {
    return this.http.get<Questions[]>(this.baseUrl1+'questions');
  }
  createPolicy(question: QuestionsNew) {
    return this.http.post(this.baseUrl1+'policy', question);
  } // Create New Insured Account

  getPolicy() {
    return this.http.get<Questions[]>(this.baseUrl1+'policy');
  } // Create New Insured Account
  
  getUsers() {
    return this.http.get<User[]>(this.baseUrl1+'userrole');
  } // Get Profile
  createProfile(user: User) {
    return this.http.post(this.baseUrl1+'userrole', user);
  } // Create New Profile
}
