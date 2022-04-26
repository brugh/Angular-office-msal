import { Component, OnInit } from '@angular/core';
import { Observable } from 'rxjs';
import { SharepointService } from '../sharepoint.service';

@Component({
  selector: 'app-sharepoint',
  templateUrl: './sharepoint.component.html',
  styleUrls: ['./sharepoint.component.scss']
})
export class SharepointComponent implements OnInit {
  sharepoint$?: Observable<File[]>;

  constructor(private sharepoint: SharepointService) { }

  ngOnInit(): void {
    this.sharepoint$ = this.sharepoint.getFiles();
  }

}
