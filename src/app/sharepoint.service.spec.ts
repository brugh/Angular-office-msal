import { TestBed } from '@angular/core/testing';

import { SharepointService } from './sharepoint.service';

describe('SharepointService', () => {
  let service: SharepointService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(SharepointService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
